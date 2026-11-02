// Genererer MIAI's ikoner uden eksterne afhængigheder:
//   assets/tray.png        menu-bar-ikon (template: sort glyph + alpha)
//   build/MIAI.iconset/    app-ikoner (16–1024 px)
//   build/icon.icns        via `iconutil` (kun macOS)
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateSync } from "node:zlib";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");

// --- PNG-encoder -------------------------------------------------------------
let CRC_TABLE;
function crc32(buf) {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const t = Buffer.from(type, "ascii");
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([t, data])));
  return Buffer.concat([len, t, data, crc]);
}

function writePng(path, size, rgba) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); // bredde
  ihdr.writeUInt32BE(size, 4); // højde
  ihdr[8] = 8;  // bitdybde
  ihdr[9] = 6;  // RGBA
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, png);
}

// --- Tegning ------------------------------------------------------------------
/** Afstand fra punkt (x,y) til linjestykke (ax,ay)-(bx,by). */
function segDist(px, py, ax, ay, bx, by) {
  const dx = bx - ax;
  const dy = by - ay;
  const l2 = dx * dx + dy * dy;
  let t = l2 ? ((px - ax) * dx + (py - ay) * dy) / l2 : 0;
  t = Math.max(0, Math.min(1, t));
  const cx = ax + t * dx;
  const cy = ay + t * dy;
  return Math.hypot(px - cx, py - cy);
}

/**
 * MIAI's "M"-glyph som tykke linjestykker i enhedsbokse (0..1).
 * Returnerer true, hvis punktet er inde i glyphen.
 */
function inM(px, py) {
  const t = 0.13; // stregtykkelse i enhedsenheder
  const A = [0.26, 0.78], B = [0.26, 0.3], C = [0.5, 0.64], D = [0.74, 0.3], E = [0.74, 0.78];
  return (
    segDist(px, py, ...A, ...B) < t / 2 || // venstre stam
    segDist(px, py, ...B, ...C) < t / 2 || // venstre diagonal
    segDist(px, py, ...C, ...D) < t / 2 || // højre diagonal
    segDist(px, py, ...D, ...E) < t / 2 // højre stam
  );
}

function lerp(a, b, t) { return a + (b - a) * t; }

/**
 * Tegner app-ikonet: afrundet firkant med indigo-gradient og hvid "M".
 * size × size RGBA-bufre. Antialiasing via smoothstep på kant-afstande.
 */
function drawAppIcon(size) {
  const buf = Buffer.alloc(size * size * 4);
  const r = size * 0.225; // radius
  const top = [79, 70, 229];    // #4F46E5
  const bottom = [124, 58, 237]; // #7C3AED
  const aa = Math.max(1, size / 512); // antialias-bredde

  const cover = (x, y) => {
    // Afrundet rektangel: SDF er negativt inde i; coverage falder over kant.
    const dx = Math.max(Math.abs(x - size / 2) - (size / 2 - r), 0);
    const dy = Math.max(Math.abs(y - size / 2) - (size / 2 - r), 0);
    return smooth(r - Math.hypot(dx, dy) + aa / 2, aa);
  };

  for (let y = 0; y < size; y++) {
    const g = lerp(0, 1, y / (size - 1));
    for (let x = 0; x < size; x++) {
      const i = (y * size + x) * 4;
      const cBg = cover(x, y);
      if (cBg <= 0) continue;
      const px = x / size;
      const py = y / size;
      // Glyph: afstand til glyph-kant i pixels.
      const inset = 0; // glyph tegnes direkte i enhedsbokse
      let cGlyph = 0;
      const t = size * 0.13 / 2 + aa / 2;
      const d = Math.min(
        segDist(px, py, 0.26 + inset, 0.78, 0.26 + inset, 0.3),
        segDist(px, py, 0.26 + inset, 0.3, 0.5, 0.64),
        segDist(px, py, 0.5, 0.64, 0.74 + inset, 0.3),
        segDist(px, py, 0.74 + inset, 0.3, 0.74 + inset, 0.78),
      );
      cGlyph = smooth(t - d * size, aa); // 1 inde i glyphen
      const w = cGlyph * cBg;
      const rCh = lerp(top[0], 255, w);
      const gCh = lerp(lerp(top[1], bottom[1], g), 255, w);
      const bCh = lerp(lerp(top[2], bottom[2], g), 255, w);
      buf[i] = rCh;
      buf[i + 1] = gCh;
      buf[i + 2] = bCh;
      buf[i + 3] = Math.round(255 * cBg);
    }
  }
  return buf;
}

function smooth(d, aa) {
  const t = Math.max(0, Math.min(1, (d + aa / 2) / aa));
  return t * t * (3 - 2 * t);
}

/** Menu-bar-ikon: transparent baggrund, sort M (template-ikon). */
function drawTray(size) {
  const buf = Buffer.alloc(size * size * 4);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const i = (y * size + x) * 4;
      const px = (x + 0.5) / size;
      const py = (y + 0.5) / size;
      // Krymp glyphen lidt, så den hænger i menuen.
      const c = inM(px * 0.86 + 0.07, py * 0.86 + 0.07) ? 1 : 0;
      buf[i] = 0;
      buf[i + 1] = 0;
      buf[i + 2] = 0;
      buf[i + 3] = c * 255;
    }
  }
  return buf;
}

/** Block-gennemsnit: skalering ned fra master-bufre. */
function downscale(master, srcSize, dstSize) {
  const out = Buffer.alloc(dstSize * dstSize * 4);
  const factor = srcSize / dstSize;
  for (let y = 0; y < dstSize; y++) {
    for (let x = 0; x < dstSize; x++) {
      const sx = Math.floor(x * factor);
      const sy = Math.floor(y * factor);
      const sw = Math.ceil((x + 1) * factor) - sx;
      const sh = Math.ceil((y + 1) * factor) - sy;
      let r = 0, g = 0, b = 0, a = 0;
      for (let dy = 0; dy < sh; dy++) {
        for (let dx = 0; dx < sw; dx++) {
          const i = ((sy + dy) * srcSize + (sx + dx)) * 4;
          r += master[i]; g += master[i + 1]; b += master[i + 2]; a += master[i + 3];
        }
      }
      const n = sw * sh;
      const o = (y * dstSize + x) * 4;
      out[o] = Math.round(r / n);
      out[o + 1] = Math.round(g / n);
      out[o + 2] = Math.round(b / n);
      out[o + 3] = Math.round(a / n);
    }
  }
  return out;
}

// --- Hoved ---------------------------------------------------------------------
const MASTER = 1024;
console.log("Tegner app-ikon (1024 px)…");
const master = drawAppIcon(MASTER);

writePng(join(root, "assets", "tray.png"), 44, drawTray(44));
console.log("✓ assets/tray.png");

const iconset = join(root, "build", "MIAI.iconset");
rmSync(iconset, { recursive: true, force: true });
mkdirSync(iconset, { recursive: true });
const targets = [
  ["icon_16x16.png", 16],
  ["icon_16x16@2x.png", 32],
  ["icon_32x32.png", 32],
  ["icon_32x32@2x.png", 64],
  ["icon_128x128.png", 128],
  ["icon_128x128@2x.png", 256],
  ["icon_256x256.png", 256],
  ["icon_256x256@2x.png", 512],
  ["icon_512x512.png", 512],
  ["icon_512x512@2x.png", 1024],
];
for (const [name, size] of targets) {
  writePng(join(iconset, name), size, size === MASTER ? master : downscale(master, MASTER, size));
}
console.log(`✓ ${iconset}`);

if (process.platform === "darwin") {
  execFileSync("iconutil", ["-c", "icns", iconset, "-o", join(root, "build", "icon.icns")]);
  console.log("✓ build/icon.icns");
} else {
  console.log("(iconutil kun på macOS – .icns springes over)");
}
