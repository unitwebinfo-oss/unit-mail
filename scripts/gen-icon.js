// Gera build/icon.png (512x512) — quadrado azul arredondado + envelope branco.
// Renderiza em 2048 e reduz 4x para bordas suaves (antialiasing simples).
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');

const S = 4;              // supersampling
const W = 512 * S, H = 512 * S;
const px = new Float32Array(W * H * 4); // acumulador RGBA

const inRoundRect = (x, y, x0, y0, x1, y1, r) => {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.max(x0 + r, Math.min(x, x1 - r));
  const cy = Math.max(y0 + r, Math.min(y, y1 - r));
  return (x - cx) * (x - cx) + (y - cy) * (y - cy) <= r * r;
};

const distSeg = (px_, py, x1, y1, x2, y2) => {
  const dx = x2 - x1, dy = y2 - y1;
  const t = Math.max(0, Math.min(1, ((px_ - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)));
  const qx = x1 + t * dx, qy = y1 + t * dy;
  return Math.hypot(px_ - qx, py - qy);
};

const BLUE = [0x0f, 0x6c, 0xbd];
const BLUE_D = [0x0b, 0x54, 0x95];
const WHITE = [0xff, 0xff, 0xff];

for (let y = 0; y < H; y++) {
  for (let x = 0; x < W; x++) {
    const i = (y * W + x) * 4;
    // Fundo: quadrado azul com cantos arredondados (raio ~22%)
    if (!inRoundRect(x, y, 0, 0, W - 1, H - 1, 460 * S / 2)) continue;
    // leve gradiente vertical azul
    const t = y / H;
    px[i] = BLUE[0] + (BLUE_D[0] - BLUE[0]) * t;
    px[i + 1] = BLUE[1] + (BLUE_D[1] - BLUE[1]) * t;
    px[i + 2] = BLUE[2] + (BLUE_D[2] - BLUE[2]) * t;
    px[i + 3] = 255;
    // Envelope: retângulo branco central (58% largura, 44% altura)
    const ex0 = 0.21 * W, ex1 = 0.79 * W, ey0 = 0.30 * H, ey1 = 0.70 * H;
    if (inRoundRect(x, y, ex0, ey0, ex1, ey1, 26 * S)) {
      px[i] = px[i + 1] = px[i + 2] = 255;
      px[i + 3] = 255;
      // Aba do envelope: V do topo até o centro (linhas azuis)
      const midX = (ex0 + ex1) / 2, apexY = ey0 + (ey1 - ey0) * 0.52;
      const lw = 9 * S;
      const d1 = distSeg(x, y, ex0, ey0, midX, apexY);
      const d2 = distSeg(x, y, ex1, ey0, midX, apexY);
      const d = Math.min(d1, d2);
      if (d < lw) {
        // blend azul sobre branco conforme distância ao centro da linha
        const a = Math.max(0, 1 - d / lw);
        px[i] = px[i] * (1 - a) + BLUE[0] * a;
        px[i + 1] = px[i + 1] * (1 - a) + BLUE[1] * a;
        px[i + 2] = px[i + 2] * (1 - a) + BLUE[2] * a;
      }
    }
  }
}

// Downsample 4x → 512
const OW = 512;
const out = Buffer.alloc(OW * OW * 4);
for (let y = 0; y < OW; y++) {
  for (let x = 0; x < OW; x++) {
    let r = 0, g = 0, b = 0, a = 0;
    for (let sy = 0; sy < S; sy++)
      for (let sx = 0; sx < S; sx++) {
        const i = ((y * S + sy) * W + x * S + sx) * 4;
        r += px[i]; g += px[i + 1]; b += px[i + 2]; a += px[i + 3];
      }
    const n = S * S, o = (y * OW + x) * 4;
    out[o] = r / n; out[o + 1] = g / n; out[o + 2] = b / n; out[o + 3] = a / n;
  }
}

// Encoder PNG mínimo (RGBA 8-bit, filtro 0)
const crcT = new Int32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  crcT[n] = c;
}
const crc32 = (buf) => {
  let c = -1;
  for (const b of buf) c = crcT[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
};
const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(OW, 0); ihdr.writeUInt32BE(OW, 4);
ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA
const raw = Buffer.alloc(OW * (OW * 4 + 1));
for (let y = 0; y < OW; y++) {
  raw[y * (OW * 4 + 1)] = 0;
  out.copy(raw, y * (OW * 4 + 1) + 1, y * OW * 4, (y + 1) * OW * 4);
}
const png = Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk('IHDR', ihdr),
  chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
  chunk('IEND', Buffer.alloc(0)),
]);
const outPath = path.join(__dirname, '..', 'build', 'icon.png');
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, png);
console.log('ok:', outPath, png.length, 'bytes');
