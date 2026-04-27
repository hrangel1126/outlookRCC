// setup.js — Generates placeholder icon PNG files in the assets/ folder
//
// Run once: node setup.js
//
// Creates simple colored square PNGs (16x16, 32x32, 64x64, 80x80, 128x128).
// Replace these with real branded icons before deploying to users.
//
// The PNG files are built from raw bytes — no npm dependencies needed.

const fs   = require("fs");
const path = require("path");
const zlib = require("zlib"); // built into Node.js

const ASSETS_DIR = path.join(__dirname, "assets");
if (!fs.existsSync(ASSETS_DIR)) fs.mkdirSync(ASSETS_DIR);

// RCC blue: #0078d4  (R=0, G=120, B=212)
const R = 0, G = 120, B = 212;

// Build a minimal valid PNG for an N×N square of a solid color
function makePNG(size) {
  // --- PNG signature ---
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // --- IHDR chunk ---
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);  // width
  ihdrData.writeUInt32BE(size, 4);  // height
  ihdrData[8]  = 8;   // bit depth
  ihdrData[9]  = 2;   // color type: RGB
  ihdrData[10] = 0;   // compression
  ihdrData[11] = 0;   // filter
  ihdrData[12] = 0;   // interlace
  const ihdr = makeChunk("IHDR", ihdrData);

  // --- IDAT chunk: raw pixel rows, each prefixed with filter byte 0 ---
  const row = Buffer.alloc(1 + size * 3); // filter(1) + RGB per pixel
  row[0] = 0; // filter type: None
  for (let x = 0; x < size; x++) {
    row[1 + x * 3]     = R;
    row[1 + x * 3 + 1] = G;
    row[1 + x * 3 + 2] = B;
  }
  const rawPixels = Buffer.concat(Array(size).fill(row));
  const compressed = zlib.deflateSync(rawPixels);
  const idat = makeChunk("IDAT", compressed);

  // --- IEND chunk ---
  const iend = makeChunk("IEND", Buffer.alloc(0));

  return Buffer.concat([sig, ihdr, idat, iend]);
}

// Wrap data in a PNG chunk: length (4) + type (4) + data + CRC (4)
function makeChunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const lenBuf    = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);

  // CRC covers type + data
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBytes, data])), 0);

  return Buffer.concat([lenBuf, typeBytes, data, crcBuf]);
}

// CRC-32 implementation (required by PNG spec)
const CRC_TABLE = (function () {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

// Generate one PNG per required size
const SIZES = { "icon-16.png": 16, "icon-32.png": 32, "icon-64.png": 64,
                "icon-80.png": 80, "icon-128.png": 128 };

Object.entries(SIZES).forEach(function ([name, size]) {
  const dest = path.join(ASSETS_DIR, name);
  fs.writeFileSync(dest, makePNG(size));
  console.log("Created: assets/" + name + " (" + size + "x" + size + " px, blue placeholder)");
});

console.log("\nDone. Replace these files with real branded icons before deploying.");
