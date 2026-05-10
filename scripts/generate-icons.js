/**
 * Generates PNG icon files for the Outlook Follow-up add-in.
 * Run with: node scripts/generate-icons.js
 *
 * Produces a blue (#0078D4) rounded rectangle with a white calendar icon
 * using only Node.js built-ins (no extra npm packages).
 */

"use strict";
const fs   = require("fs");
const path = require("path");
const zlib = require("zlib");

// ── CRC32 ──────────────────────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) crc = CRC_TABLE[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function chunk(type, data) {
  const tBuf = Buffer.from(type, "ascii");
  const len  = Buffer.allocUnsafe(4); len.writeUInt32BE(data.length, 0);
  const crcv = crc32(Buffer.concat([tBuf, data]));
  const cBuf = Buffer.allocUnsafe(4); cBuf.writeUInt32BE(crcv, 0);
  return Buffer.concat([len, tBuf, data, cBuf]);
}

// ── PNG builder (RGBA) ─────────────────────────────────────────────────────
function buildPNG(pixels, size) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  const ihdrData = Buffer.allocUnsafe(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8]  = 8; // bit depth
  ihdrData[9]  = 6; // color type: RGBA
  ihdrData[10] = ihdrData[11] = ihdrData[12] = 0;

  const rowLen = 1 + size * 4;
  const raw    = Buffer.alloc(rowLen * size);
  for (let y = 0; y < size; y++) {
    raw[y * rowLen] = 0; // filter: None
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixels[y][x];
      const off = y * rowLen + 1 + x * 4;
      raw[off] = r; raw[off+1] = g; raw[off+2] = b; raw[off+3] = a;
    }
  }

  return Buffer.concat([
    sig,
    chunk("IHDR", ihdrData),
    chunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// ── Icon design (scales to any size) ─────────────────────────────────────
// Blue bg + white calendar shape
const BLUE  = [0x00, 0x78, 0xD4, 0xFF];
const WHITE = [0xFF, 0xFF, 0xFF, 0xFF];
const TRANS = [0x00, 0x00, 0x00, 0x00];

function makeIconPixels(size) {
  const pixels = Array.from({ length: size }, () =>
    Array.from({ length: size }, () => [...BLUE])
  );

  // Rounded corners (remove 2-pixel corner boxes)
  const r = 2;
  for (let y = 0; y < r; y++) {
    for (let x = 0; x < r; x++) {
      if (x + y < r) {
        pixels[y][x] = TRANS; pixels[y][size-1-x] = TRANS;
        pixels[size-1-y][x] = TRANS; pixels[size-1-y][size-1-x] = TRANS;
      }
    }
  }

  // Helper: fill rectangle with color
  const fill = (x1, y1, x2, y2, col) => {
    for (let y = y1; y <= y2; y++)
      for (let x = x1; x <= x2; x++)
        if (x >= 0 && x < size && y >= 0 && y < size)
          pixels[y][x] = col;
  };

  // Scale factors relative to a 80px reference
  const s = size / 80;
  const p = (v) => Math.round(v * s); // scale a pixel value

  // Calendar body outline (white rounded rect, then blue interior)
  fill(p(10), p(14), p(70), p(70), WHITE);
  fill(p(13), p(17), p(67), p(67), BLUE);

  // Top bar (white, darker)
  fill(p(10), p(14), p(70), p(28), WHITE);

  // Calendar tabs (blue cutouts on the top bar)
  fill(p(20), p(10), p(30), p(22), WHITE);
  fill(p(50), p(10), p(60), p(22), WHITE);
  fill(p(22), p(12), p(28), p(24), BLUE);
  fill(p(52), p(12), p(58), p(24), BLUE);

  // Grid lines (white) — 3 rows × 4 cols
  const col1 = p(10), col5 = p(70);
  for (let row = 0; row < 3; row++) {
    const ty = p(34 + row * 12);
    fill(col1+1, ty, col5-1, ty+1, WHITE);
  }
  for (let col = 0; col < 3; col++) {
    const tx = p(25 + col * 15);
    fill(tx, p(29), tx+1, p(70), WHITE);
  }

  return pixels;
}

// ── Write files ────────────────────────────────────────────────────────────
const assetsDir = path.resolve(__dirname, "..", "assets");
if (!fs.existsSync(assetsDir)) fs.mkdirSync(assetsDir, { recursive: true });

for (const size of [16, 32, 80]) {
  const buf = buildPNG(makeIconPixels(size), size);
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), buf);
  console.log(`✓  icon-${size}.png  (${buf.length} bytes)`);
}

// 192×192 colour icon for Teams/Unified manifest
const buf192 = buildPNG(makeIconPixels(192), 192);
fs.writeFileSync(path.join(assetsDir, "icon-color.png"), buf192);
console.log(`✓  icon-color.png  (192×192, ${buf192.length} bytes)`);

// 32×32 white-on-transparent outline icon
const outlinePixels = Array.from({ length: 32 }, (_, y) =>
  Array.from({ length: 32 }, (_, x) => {
    const base = makeIconPixels(32)[y][x];
    // Swap blue→transparent, white→white
    return (base[3] === 0) ? TRANS
         : (base[0] === 0xFF) ? WHITE
         : TRANS;
  })
);
const bufOutline = buildPNG(outlinePixels, 32);
fs.writeFileSync(path.join(assetsDir, "icon-outline.png"), bufOutline);
console.log(`✓  icon-outline.png  (${bufOutline.length} bytes)`);

console.log("\nDone! All icons written to assets/");
