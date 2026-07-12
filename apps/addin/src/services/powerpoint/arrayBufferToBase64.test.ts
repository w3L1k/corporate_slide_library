import { describe, expect, it } from "vitest";
import { arrayBufferToBase64 } from "./arrayBufferToBase64";

describe("arrayBufferToBase64", () => {
  it("encodes empty and known binary values", () => {
    expect(arrayBufferToBase64(new ArrayBuffer(0))).toBe("");
    expect(arrayBufferToBase64(new Uint8Array([0, 1, 2, 3, 254, 255]).buffer)).toBe(
      "AAECA/7/"
    );
  });

  it("preserves data across multiple encoding chunks", () => {
    const source = Uint8Array.from({ length: 70_003 }, (_, index) => index % 251);
    const encoded = arrayBufferToBase64(source.buffer);
    const decoded = Uint8Array.from(atob(encoded), (character) => character.charCodeAt(0));

    expect(decoded).toEqual(source);
  });
});
