// Genererer simple PNG-ikoner til manifestet uden eksterne afhængigheder.
// Kør: node scripts/make-icons.mjs
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync } from "node:zlib";

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const c = Buffer.alloc(4);
  c.writeUInt32BE(crc(td));
  return Buffer.concat([len, td, c]);
};

function icon(size) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  const r = size * 0.22; // hjørneradius
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      const i = y * (size * 4 + 1) + 1 + x * 4;
      // afrundet firkant
      const dx = Math.max(r - x - 0.5, x + 0.5 - (size - r), 0);
      const dy = Math.max(r - y - 0.5, y + 0.5 - (size - r), 0);
      const inside = dx * dx + dy * dy <= r * r;
      // stiliseret "chip": hvid ring med prik i midten
      const cx = x + 0.5 - size / 2,
        cy = y + 0.5 - size / 2,
        d = Math.hypot(cx, cy);
      const ring = d > size * 0.22 && d < size * 0.32;
      const dot = d < size * 0.1;
      const [R, G, B] = ring || dot ? [255, 255, 255] : [15, 108, 189];
      raw[i] = R;
      raw[i + 1] = G;
      raw[i + 2] = B;
      raw[i + 3] = inside ? 255 : 0;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}

const out = new URL("../packages/addin/public/assets/", import.meta.url);
mkdirSync(out, { recursive: true });
for (const s of [16, 32, 64, 80, 128]) writeFileSync(new URL(`icon-${s}.png`, out), icon(s));
console.log("Ikoner skrevet til packages/addin/public/assets");
