/**
 * setup.js
 * Generates PNG icon files and installs office-addin-dev-certs for HTTPS.
 * Run once before starting the server: node setup.js
 */

const fs   = require("fs");
const path = require("path");
const zlib = require("zlib");
const { execSync } = require("child_process");

// ─── PNG generator (no external dependencies) ────────────────────────────────

function makeCRCTable() {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
}
const CRC_TABLE = makeCRCTable();

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buf[i]) & 0xff];
  crc = (crc ^ 0xffffffff) >>> 0;
  const out = Buffer.alloc(4);
  out.writeUInt32BE(crc);
  return out;
}

function makeChunk(type, data) {
  const len  = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const typeB = Buffer.from(type, "ascii");
  return Buffer.concat([len, typeB, data, crc32(Buffer.concat([typeB, data]))]);
}

/**
 * Generate a simple square PNG icon.
 * Draws a blue rounded background with a white triangle (▲) using pixel math.
 */
function generateIconPNG(size) {
  // Colours
  const BG_R = 0x00, BG_G = 0x78, BG_B = 0xd4; // #0078d4 (Microsoft blue)
  const FG_R = 0xff, FG_G = 0xff, FG_B = 0xff; // white

  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  // IHDR
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 2;  // colour type: RGB

  // Raw pixel data
  const raw = Buffer.alloc((size * 3 + 1) * size);
  let offset = 0;

  // Triangle vertices (centred, occupying ~55% of the icon)
  const margin  = Math.round(size * 0.22);
  const tipX    = size / 2;
  const tipY    = margin;
  const baseY   = size - margin;
  const baseL   = margin;
  const baseR   = size - margin;

  for (let y = 0; y < size; y++) {
    raw[offset++] = 0; // filter byte = None
    for (let x = 0; x < size; x++) {
      // Simple point-in-triangle test
      const inTriangle =
        y >= tipY && y <= baseY &&
        x >= baseL + ((tipX - baseL) * (y - baseY)) / (tipY - baseY) &&
        x <= baseR + ((tipX - baseR) * (y - baseY)) / (tipY - baseY);

      if (inTriangle) {
        raw[offset++] = FG_R;
        raw[offset++] = FG_G;
        raw[offset++] = FG_B;
      } else {
        raw[offset++] = BG_R;
        raw[offset++] = BG_G;
        raw[offset++] = BG_B;
      }
    }
  }

  const compressed = zlib.deflateSync(raw);

  return Buffer.concat([
    signature,
    makeChunk("IHDR", ihdr),
    makeChunk("IDAT", compressed),
    makeChunk("IEND", Buffer.alloc(0)),
  ]);
}

// ─── Create icons ────────────────────────────────────────────────────────────

const assetsDir = path.join(__dirname, "src", "assets");
fs.mkdirSync(assetsDir, { recursive: true });

const sizes = [16, 32, 80];
for (const sz of sizes) {
  const dest = path.join(assetsDir, `icon-${sz}.png`);
  fs.writeFileSync(dest, generateIconPNG(sz));
  console.log(`  Created ${dest}`);
}
console.log("Icons generated.\n");

// ─── Install dev certs ───────────────────────────────────────────────────────

console.log("Installing office-addin-dev-certs for HTTPS (may require elevated prompt)…");
try {
  execSync("npx office-addin-dev-certs install --days 3650", { stdio: "inherit" });
  console.log("\nDev certs installed. You can now run: npm start\n");
} catch (err) {
  console.warn(
    "\nWarning: Could not install dev certs automatically.\n" +
    "Run manually:  npx office-addin-dev-certs install\n" +
    "Or install globally: npm i -g office-addin-dev-certs && office-addin-dev-certs install\n"
  );
}
