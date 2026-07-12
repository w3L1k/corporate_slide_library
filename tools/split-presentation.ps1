[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePptx,

    [Parameter(Mandatory = $true)]
    [string]$LibraryRoot,

    [string]$Category = "Imported",
    [string]$Owner = "Content owner",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

function Get-SafeId {
    param([string]$Value, [int]$Index)

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormKD).ToLowerInvariant()
    $slug = [Text.RegularExpressions.Regex]::Replace($normalized, "[^a-z0-9]+", "-").Trim("-")
    if ([string]::IsNullOrWhiteSpace($slug)) {
        $slug = "slide-$($Index.ToString('000'))"
    }
    return "$slug-$($Index.ToString('000'))"
}

function Get-SlideTitle {
    param($Slide, [int]$Index)

    try {
        if ($null -ne $Slide.Shapes.Title -and $Slide.Shapes.Title.HasTextFrame -eq -1) {
            $candidate = $Slide.Shapes.Title.TextFrame.TextRange.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                return $candidate
            }
        }
    }
    catch {
        # Some layouts have no title placeholder; use the stable fallback below.
    }
    return "Imported slide $Index"
}

$sourcePath = [IO.Path]::GetFullPath($SourcePptx)
$libraryPath = [IO.Path]::GetFullPath($LibraryRoot)
if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
    throw "Source presentation does not exist: $sourcePath"
}
if ([IO.Path]::GetExtension($sourcePath) -ne ".pptx") {
    throw "Source presentation must use the .pptx extension."
}

$slidesPath = Join-Path $libraryPath "slides"
$previewsPath = Join-Path $libraryPath "previews"
New-Item -ItemType Directory -Force -Path $slidesPath, $previewsPath | Out-Null

$powerPoint = $null
$source = $null
$items = [Collections.Generic.List[object]]::new()

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $source = $powerPoint.Presentations.Open($sourcePath, $true, $false, $false)

    for ($index = 1; $index -le $source.Slides.Count; $index++) {
        $slide = $source.Slides.Item($index)
        $title = Get-SlideTitle -Slide $slide -Index $index
        $id = Get-SafeId -Value $title -Index $index
        $slideFile = Join-Path $slidesPath "$id.pptx"
        $previewFile = Join-Path $previewsPath "$id.png"

        if (-not $Overwrite -and ((Test-Path -LiteralPath $slideFile) -or (Test-Path -LiteralPath $previewFile))) {
            throw "Output already exists for $id. Use -Overwrite to replace it."
        }

        $source.SaveCopyAs($slideFile, 24, 0)
        $singleSlide = $powerPoint.Presentations.Open($slideFile, $false, $false, $false)
        try {
            for ($deleteIndex = $singleSlide.Slides.Count; $deleteIndex -ge 1; $deleteIndex--) {
                if ($deleteIndex -ne $index) {
                    $singleSlide.Slides.Item($deleteIndex).Delete()
                }
            }
            $singleSlide.Save()
        }
        finally {
            $singleSlide.Close()
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($singleSlide) | Out-Null
        }

        $slide.Export($previewFile, "PNG", 1280, 720)
        $items.Add([ordered]@{
            id = $id
            title = $title
            description = "Imported from $([IO.Path]::GetFileName($sourcePath)), slide $index."
            category = $Category
            tags = @("imported", $Category.ToLowerInvariant())
            version = "1.0"
            status = "draft"
            updatedAt = [DateTimeOffset]::Now.ToString("o")
            sourceFile = "slides/$id.pptx"
            previewFile = "previews/$id.png"
            owner = $Owner
        })
    }

    $metadataPath = Join-Path $libraryPath "catalog.imported.json"
    ConvertTo-Json -InputObject $items.ToArray() -Depth 6 | Set-Content -LiteralPath $metadataPath -Encoding UTF8
    Write-Host "Created $($items.Count) one-slide presentations and previews."
    Write-Host "Review and merge metadata from: $metadataPath"
}
finally {
    if ($null -ne $source) {
        $source.Close()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($source) | Out-Null
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
