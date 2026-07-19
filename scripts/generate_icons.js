// ============================================================
// PWA用アイコン生成スクリプト
//
//   外部ライブラリなし（Node組み込みの zlib のみ）で、
//   藍色の背景 + 白いブックマーク型のアイコンPNGを生成する。
//   maskable対応のため、図柄は中央80%のセーフゾーンに収める。
//
//   実行方法: node scripts/generate_icons.js
//   出力先  : icons/icon-192.png, icons/icon-512.png
// ============================================================

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');

// 背景色（藍色）とブックマークの色（白）
const BG_COLOR = [0x39, 0x49, 0xab];
const MARK_COLOR = [0xff, 0xff, 0xff];

/**
 * PNGチャンク（長さ + タイプ + データ + CRC32）を組み立てる
 */
function buildChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcInput = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  // zlib.crc32 は符号付きになりうるため >>> 0 で符号なしに揃える
  crc.writeUInt32BE(zlib.crc32(crcInput) >>> 0);
  return Buffer.concat([length, typeBuf, data, crc]);
}

/**
 * RGBAピクセル配列からPNGバイナリを組み立てる
 */
function encodePng(size, pixels) {
  // IHDR: 幅・高さ・ビット深度8・カラータイプ6(RGBA)
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // ビット深度
  ihdr[9] = 6; // RGBA
  // ihdr[10..12] = 圧縮方式・フィルタ方式・インターレース（すべて0）

  // 各行の先頭にフィルタタイプ0（None）を付けてdeflate圧縮
  const stride = size * 4;
  const raw = Buffer.alloc((stride + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (stride + 1)] = 0; // フィルタ: None
    pixels.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
  }
  const idat = zlib.deflateSync(raw, { level: 9 });

  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    buildChunk('IHDR', ihdr),
    buildChunk('IDAT', idat),
    buildChunk('IEND', Buffer.alloc(0))
  ]);
}

/**
 * 指定サイズのアイコンピクセルを描画する。
 * 背景は全面塗り（maskable対応）、中央にブックマーク型の白図形。
 */
function drawIcon(size) {
  const pixels = Buffer.alloc(size * size * 4);

  // ブックマークの形状パラメータ（セーフゾーン: 中央80%に収まる比率）
  const markWidth = size * 0.34;   // ブックマークの幅
  const markTop = size * 0.26;     // 上端
  const markBottom = size * 0.74;  // 下端
  const notchDepth = size * 0.10;  // 下端中央の切り込みの深さ
  const centerX = size / 2;
  const left = centerX - markWidth / 2;
  const right = centerX + markWidth / 2;

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let color = BG_COLOR;

      // ブックマーク判定: 横範囲内かつ、V字切り込みを除いた縦範囲内
      if (x >= left && x <= right && y >= markTop) {
        // 下端はV字: 中央ほど浅く切れ込む（|x-centerX| が小さいほど上に上がる）
        const ratio = Math.abs(x - centerX) / (markWidth / 2);
        const bottomEdge = markBottom - notchDepth * (1 - ratio);
        if (y <= bottomEdge) {
          color = MARK_COLOR;
        }
      }

      const offset = (y * size + x) * 4;
      pixels[offset] = color[0];
      pixels[offset + 1] = color[1];
      pixels[offset + 2] = color[2];
      pixels[offset + 3] = 255; // 不透明
    }
  }
  return pixels;
}

// ---- メイン処理 ----------------------------------------------

const outDir = path.join(__dirname, '..', 'icons');
fs.mkdirSync(outDir, { recursive: true });

for (const size of [192, 512]) {
  const png = encodePng(size, drawIcon(size));
  const outPath = path.join(outDir, `icon-${size}.png`);
  fs.writeFileSync(outPath, png);
  console.log(`生成完了: ${outPath} (${png.length} bytes)`);
}
