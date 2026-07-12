// Divisible by three so independently encoded chunks can be concatenated safely.
const BYTE_CHUNK_SIZE = 24_576;

/** Converts binary data without spreading the full byte array onto the call stack. */
export const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  const base64Chunks: string[] = [];

  for (let offset = 0; offset < bytes.length; offset += BYTE_CHUNK_SIZE) {
    const limit = Math.min(offset + BYTE_CHUNK_SIZE, bytes.length);
    let binaryChunk = "";

    for (let index = offset; index < limit; index += 1) {
      binaryChunk += String.fromCharCode(bytes[index] ?? 0);
    }

    base64Chunks.push(btoa(binaryChunk));
  }

  return base64Chunks.join("");
};
