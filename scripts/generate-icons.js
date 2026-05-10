/**
 * Generates PNG icon files for the Outlook Follow-up add-in.
 * Run with: node scripts/generate-icons.js
 *
 * Design: bold white "F" lettermark on Microsoft-blue (#0078D4) background,
 * with rounded corners. Readable at all sizes from 16x16 upward.
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
  const cBuf = Buffer.allocUnsafe(4); cBuf.writeUInt32BE(crc32(Buffer.concat([tBuf, data])), 0);
  return Buffer.concat([len, tBuf, data, cBuf]);
}

// ── PNG builder (RGBA) ─────────────────────────────────────────────────────
function buildPNG(grid, size) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.allocUnsafe(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = ihdr[11] = ihdr[12] = 0;

  const rowLen = 1 + size * 4;
  const raw    = Buffer.alloc(rowLen * size);
  for (let y = 0; y < size; y++) {
    raw[y * rowLen] = 0;
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = grid[y * size + x];
      const o = y * rowLen + 1 + x * 4;
      raw[o] = r; raw[o+1] = g; raw[o+2] = b; raw[o+3] = a;
    }
  }
  return Buffer.concat([sig, chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0))]);
}

// ── Pixel helpers ──────────────────────────────────────────────────────────
const BLUE  = [0x00, 0x78, 0xD4, 0xFF];
const WHITE = [0xFF, 0xFF, 0xFF, 0xFF];
const TRANS = [0x00, 0x00, 0x00, 0x00];

function makeGrid(size) {
  return new Array(size * size).fill(null).map(() => [...BLUE]);
}
function setPixel(grid, size, x, y, col) {
  if (x >= 0 && x < size && y >= 0 && y < size)
    grid[y * size + x] = col;
}
function fillRect(grid, size, x1, y1, x2, y2, col) {
  for (let y = y1; y <= y2; y++)
    for (let x = x1; x <= x2; x++)
      setPixel(grid, size, x, y, col);
}

// ── "F" lettermark design ──────────────────────────────────────────────────
function makeIcon(size) {
  const grid = makeGrid(size);
  const p = v => Math.round(v * size);

  // Rounded corners
  const r = Math.max(2, p(0.1));
  for (let d = 0; d < r; d++) {
    setPixel(grid, size,       d,       r-1-d, TRANS);
    setPixel(grid, size, size-1-d,       r-1-d, TRANS);
    setPixel(grid, size,       d, size-1-(r-1-d), TRANS);
    setPixel(grid, size, size-1-d, size-1-(r-1-d), TRANS);
  }

  // Bold "F" lettermark
  const left   = p(0.20);
  const right  = p(0.75);
  const mid    = p(0.60);
  const top    = p(0.18);
  const bot    = p(0.82);
  const midY1  = p(0.44);
  const midY2  = p(0.56);
  const stroke = Math.max(2, p(0.15));

  fillRect(grid, size, left, top, left + stroke - 1, bot, WHITE);   // vertical
  fillRect(grid, size, left, top, right, top + stroke - 1, WHITE);  // top bar
  fillRect(grid, size, left, midY1, mid, midY2, WHITE);              // middle bar

  return grid;
}

// ── Outline icon (white-on-transparent, for Teams manifest) ───────────────
function makeOutline(size) {
  return makeIcon(size).map(px =>
    px[3] === 0   ? [...TRANS] :
    px[0] === 0xFF ? [...WHITE] :
    [...TRANS]
  );
}

// ── Write files ────────────────────────────────────────────────────────────
const assetsDir = path.resolve(__dirname, "..", "assets");
if (!fs.existsSync(assetsDir)) fs.mkdirSync(assetsDir, { recursive: true });

for (const size of [16, 32, 80]) {
  const buf = buildPNG(makeIcon(size), size);
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), buf);
  console.log(`  icon-${size}.png  (${buf.length} B)`);
}

const buf192 = buildPNG(makeIcon(192), 192);
fs.writeFileSync(path.join(assetsDir, "icon-color.png"), buf192);
console.log(`  icon-color.png  192x192  (${buf192.length} B)`);

const bufOut = buildPNG(makeOutline(32), 32);
fs.writeFileSync(path.join(assetsDir, "icon-outline.png"), bufOut);
console.log(`  icon-outline.png  32x32 white-on-transparent  (${bufOut.length} B)`);

console.log("\nDone.");
