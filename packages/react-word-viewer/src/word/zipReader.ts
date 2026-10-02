/**
 * Custom browser-native ZIP archive reader without any external npm packages.
 * Uses standard DataView, Uint8Array, and browser native DecompressionStream('deflate-raw').
 * Includes a pure TypeScript fallback inflate implementation for full resilience.
 */

export interface ZipEntry {
  filename: string;
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
  crc32: number;
}

export class ZipReader {
  private buffer: ArrayBuffer;
  private view: DataView;
  private entries: Map<string, ZipEntry> = new Map();

  constructor(buffer: ArrayBuffer) {
    this.buffer = buffer;
    this.view = new DataView(buffer);
  }

  public async init(): Promise<void> {
    this.parseCentralDirectory();
  }

  public listFiles(): string[] {
    return Array.from(this.entries.keys());
  }

  public hasFile(filename: string): boolean {
    const normalized = filename.replace(/\\/g, '/');
    return this.entries.has(normalized);
  }

  public async getFile(filename: string): Promise<Uint8Array | null> {
    const normalized = filename.replace(/\\/g, '/');
    const entry = this.entries.get(normalized);
    if (!entry) {
      // Try case-insensitive search
      for (const [key, val] of this.entries.entries()) {
        if (key.toLowerCase() === normalized.toLowerCase()) {
          return this.extractEntry(val);
        }
      }
      return null;
    }
    return this.extractEntry(entry);
  }

  public async getText(filename: string, encoding: string = 'utf-8'): Promise<string | null> {
    const data = await this.getFile(filename);
    if (!data) return null;
    const decoder = new TextDecoder(encoding);
    return decoder.decode(data);
  }

  public async getBlobUrl(filename: string, mimeType?: string): Promise<string | null> {
    const data = await this.getFile(filename);
    if (!data) return null;
    const determinedMime = mimeType || this.guessMimeType(filename);
    const blob = new Blob([data as unknown as BlobPart], { type: determinedMime });
    return URL.createObjectURL(blob);
  }

  private guessMimeType(filename: string): string {
    const lower = filename.toLowerCase();
    if (lower.endsWith('.png')) return 'image/png';
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
    if (lower.endsWith('.gif')) return 'image/gif';
    if (lower.endsWith('.svg')) return 'image/svg+xml';
    if (lower.endsWith('.webp')) return 'image/webp';
    if (lower.endsWith('.xml')) return 'application/xml';
    return 'application/octet-stream';
  }

  private parseCentralDirectory(): void {
    const byteLength = this.buffer.byteLength;
    if (byteLength < 22) {
      throw new Error('Invalid file: too small to be a ZIP archive');
    }

    // Find End of Central Directory (EOCD) record
    // EOCD signature is 0x06054b50 (PK\x05\x06)
    // It is located in the last 65557 bytes of the file
    let eocdOffset = -1;
    const maxSearch = Math.min(byteLength, 65557);
    for (let i = byteLength - 22; i >= byteLength - maxSearch; i--) {
      if (this.view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }

    if (eocdOffset === -1) {
      // Fallback: Try scanning local headers from offset 0
      this.scanLocalHeaders();
      return;
    }

    const totalEntries = this.view.getUint16(eocdOffset + 10, true);
    const cdSize = this.view.getUint32(eocdOffset + 12, true);
    const cdOffset = this.view.getUint32(eocdOffset + 16, true);

    let offset = cdOffset;
    const decoder = new TextDecoder('utf-8');

    for (let i = 0; i < totalEntries && offset < cdOffset + cdSize; i++) {
      if (offset + 46 > byteLength) break;
      const signature = this.view.getUint32(offset, true);
      if (signature !== 0x02014b50) {
        // Not a central directory header
        break;
      }

      const compressionMethod = this.view.getUint16(offset + 10, true);
      const crc32 = this.view.getUint32(offset + 16, true);
      const compressedSize = this.view.getUint32(offset + 20, true);
      const uncompressedSize = this.view.getUint32(offset + 24, true);
      const filenameLength = this.view.getUint16(offset + 28, true);
      const extraLength = this.view.getUint16(offset + 30, true);
      const commentLength = this.view.getUint16(offset + 32, true);
      const localHeaderOffset = this.view.getUint32(offset + 42, true);

      const filenameBytes = new Uint8Array(this.buffer, offset + 46, filenameLength);
      const filename = decoder.decode(filenameBytes).replace(/\\/g, '/');

      this.entries.set(filename, {
        filename,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        localHeaderOffset,
        crc32,
      });

      offset += 46 + filenameLength + extraLength + commentLength;
    }
  }

  private scanLocalHeaders(): void {
    let offset = 0;
    const byteLength = this.buffer.byteLength;
    const decoder = new TextDecoder('utf-8');

    while (offset + 30 <= byteLength) {
      const signature = this.view.getUint32(offset, true);
      if (signature !== 0x04034b50) break;

      const compressionMethod = this.view.getUint16(offset + 8, true);
      const crc32 = this.view.getUint32(offset + 14, true);
      const compressedSize = this.view.getUint32(offset + 18, true);
      const uncompressedSize = this.view.getUint32(offset + 22, true);
      const filenameLength = this.view.getUint16(offset + 26, true);
      const extraLength = this.view.getUint16(offset + 28, true);

      const filenameBytes = new Uint8Array(this.buffer, offset + 30, filenameLength);
      const filename = decoder.decode(filenameBytes).replace(/\\/g, '/');

      this.entries.set(filename, {
        filename,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        localHeaderOffset: offset,
        crc32,
      });

      offset += 30 + filenameLength + extraLength + compressedSize;
    }
  }

  private async extractEntry(entry: ZipEntry): Promise<Uint8Array> {
    const localOffset = entry.localHeaderOffset;
    if (localOffset + 30 > this.buffer.byteLength) {
      throw new Error(`Invalid local header offset for ${entry.filename}`);
    }

    const localSig = this.view.getUint32(localOffset, true);
    if (localSig !== 0x04034b50) {
      throw new Error(`Invalid local header signature for ${entry.filename}`);
    }

    const filenameLength = this.view.getUint16(localOffset + 26, true);
    const extraLength = this.view.getUint16(localOffset + 28, true);
    const dataOffset = localOffset + 30 + filenameLength + extraLength;

    if (dataOffset + entry.compressedSize > this.buffer.byteLength) {
      throw new Error(`Unexpected end of data for ${entry.filename}`);
    }

    const compressedBytes = new Uint8Array(this.buffer, dataOffset, entry.compressedSize);

    // Compression method 0 = Stored (no compression)
    if (entry.compressionMethod === 0) {
      return compressedBytes;
    }

    // Compression method 8 = Deflate
    if (entry.compressionMethod === 8) {
      return this.decompressDeflate(compressedBytes, entry.uncompressedSize);
    }

    throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
  }

  private async decompressDeflate(compressedBytes: Uint8Array, expectedSize: number): Promise<Uint8Array> {
    // 1. Try standard browser DecompressionStream ('deflate-raw')
    if (typeof DecompressionStream !== 'undefined') {
      try {
        const stream = new Blob([compressedBytes as unknown as BlobPart])
          .stream()
          .pipeThrough(new DecompressionStream('deflate-raw'));
        const response = new Response(stream);
        const arrayBuf = await response.arrayBuffer();
        return new Uint8Array(arrayBuf);
      } catch {
        // Fall back to pure JS inflate if stream fails (e.g., in some edge test runners)
      }
    }

    // 2. Pure JS Inflate Fallback
    return inflateRaw(compressedBytes, expectedSize);
  }
}

/**
 * Pure JavaScript Inflate (RFC 1951 Deflate decompression) fallback implementation.
 * Zero external dependencies.
 */
function inflateRaw(input: Uint8Array, expectedSize: number): Uint8Array {
  let bitPos = 0;
  let bytePos = 0;

  function readBits(numBits: number): number {
    let result = 0;
    for (let i = 0; i < numBits; i++) {
      if (bytePos >= input.length) return result;
      const bit = (input[bytePos] >> bitPos) & 1;
      result |= bit << i;
      bitPos++;
      if (bitPos === 8) {
        bitPos = 0;
        bytePos++;
      }
    }
    return result;
  }

  const output: number[] = [];
  const maxOut = expectedSize > 0 ? expectedSize + 1024 : 10 * 1024 * 1024;

  const FIXED_LIT_LENS = new Uint8Array(288);
  for (let i = 0; i <= 143; i++) FIXED_LIT_LENS[i] = 8;
  for (let i = 144; i <= 255; i++) FIXED_LIT_LENS[i] = 9;
  for (let i = 256; i <= 279; i++) FIXED_LIT_LENS[i] = 7;
  for (let i = 280; i <= 287; i++) FIXED_LIT_LENS[i] = 8;

  const FIXED_DIST_LENS = new Uint8Array(32);
  for (let i = 0; i < 32; i++) FIXED_DIST_LENS[i] = 5;

  interface HuffmanTree {
    table: number[];
    maxBits: number;
  }

  function buildHuffman(lengths: Uint8Array | number[]): HuffmanTree {
    let maxBits = 0;
    for (let i = 0; i < lengths.length; i++) {
      if (lengths[i] > maxBits) maxBits = lengths[i];
    }
    if (maxBits === 0) return { table: [], maxBits: 0 };

    const blCount = new Uint16Array(maxBits + 1);
    for (let i = 0; i < lengths.length; i++) {
      if (lengths[i] > 0) blCount[lengths[i]]++;
    }

    const nextCode = new Uint16Array(maxBits + 1);
    let code = 0;
    for (let bits = 1; bits <= maxBits; bits++) {
      code = (code + blCount[bits - 1]) << 1;
      nextCode[bits] = code;
    }

    const tableSize = 1 << maxBits;
    const table = new Int32Array(tableSize);
    table.fill(-1);

    for (let i = 0; i < lengths.length; i++) {
      const len = lengths[i];
      if (len !== 0) {
        let c = nextCode[len]++;
        let rev = 0;
        for (let j = 0; j < len; j++) {
          rev |= ((c >> j) & 1) << (len - 1 - j);
        }
        const step = 1 << len;
        for (let j = rev; j < tableSize; j += step) {
          table[j] = (i << 4) | len;
        }
      }
    }

    return { table: Array.from(table), maxBits };
  }

  function decodeSymbol(tree: HuffmanTree): number {
    if (tree.maxBits === 0) return 0;
    let code = 0;
    for (let i = 0; i < tree.maxBits; i++) {
      if (bytePos >= input.length) return 256; // stop symbol
      const bit = (input[bytePos] >> bitPos) & 1;
      code |= bit << i;
      bitPos++;
      if (bitPos === 8) {
        bitPos = 0;
        bytePos++;
      }
      const entry = tree.table[code];
      if (entry !== -1 && (entry & 0x0f) === i + 1) {
        return entry >> 4;
      }
    }
    return 256;
  }

  const LENGTH_BASE = [
    3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115,
    131, 163, 195, 227, 258,
  ];
  const LENGTH_EXTRA = [
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
  ];
  const DIST_BASE = [
    1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537,
    2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
  ];
  const DIST_EXTRA = [
    0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12,
    13, 13,
  ];
  const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  let isFinal = 0;
  while (!isFinal && bytePos < input.length && output.length < maxOut) {
    isFinal = readBits(1);
    const blockType = readBits(2);

    if (blockType === 0) {
      // Uncompressed block
      bitPos = 0;
      if (bytePos < input.length) bytePos++;
      if (bytePos + 4 > input.length) break;
      const len = input[bytePos] | (input[bytePos + 1] << 8);
      bytePos += 4; // skip len and nlen
      for (let i = 0; i < len && bytePos < input.length; i++) {
        output.push(input[bytePos++]);
      }
    } else if (blockType === 1 || blockType === 2) {
      let litTree: HuffmanTree;
      let distTree: HuffmanTree;

      if (blockType === 1) {
        litTree = buildHuffman(FIXED_LIT_LENS);
        distTree = buildHuffman(FIXED_DIST_LENS);
      } else {
        const hlit = readBits(5) + 257;
        const hdist = readBits(5) + 1;
        const hclen = readBits(4) + 4;

        const clLengths = new Uint8Array(19);
        for (let i = 0; i < hclen; i++) {
          clLengths[CL_ORDER[i]] = readBits(3);
        }
        const clTree = buildHuffman(clLengths);

        const lengths = new Uint8Array(hlit + hdist);
        let idx = 0;
        while (idx < hlit + hdist) {
          const sym = decodeSymbol(clTree);
          if (sym < 16) {
            lengths[idx++] = sym;
          } else if (sym === 16) {
            const repeat = readBits(2) + 3;
            const prev = idx > 0 ? lengths[idx - 1] : 0;
            for (let r = 0; r < repeat && idx < hlit + hdist; r++) lengths[idx++] = prev;
          } else if (sym === 17) {
            const repeat = readBits(3) + 3;
            for (let r = 0; r < repeat && idx < hlit + hdist; r++) lengths[idx++] = 0;
          } else if (sym === 18) {
            const repeat = readBits(7) + 11;
            for (let r = 0; r < repeat && idx < hlit + hdist; r++) lengths[idx++] = 0;
          }
        }

        litTree = buildHuffman(lengths.subarray(0, hlit));
        distTree = buildHuffman(lengths.subarray(hlit));
      }

      while (output.length < maxOut) {
        const sym = decodeSymbol(litTree);
        if (sym === 256) break;
        if (sym < 256) {
          output.push(sym);
        } else {
          const lenIdx = sym - 257;
          const length = LENGTH_BASE[lenIdx] + readBits(LENGTH_EXTRA[lenIdx]);
          const distSym = decodeSymbol(distTree);
          const distance = DIST_BASE[distSym] + readBits(DIST_EXTRA[distSym]);

          let src = output.length - distance;
          for (let k = 0; k < length; k++) {
            output.push(output[src + k]);
          }
        }
      }
    } else {
      break;
    }
  }

  return new Uint8Array(output);
}
