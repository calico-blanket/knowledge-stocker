// ============================================================
// 開発用ローカルサーバ（外部ライブラリなし）
//
//   PWAの静的ファイルをローカル配信して動作確認するための簡易サーバ。
//   実行方法: node scripts/dev_server.js [ポート番号]（既定: 8765）
//   ブラウザで http://localhost:8765/ を開く。
//   共有シート経由の起動を模擬するには:
//     http://localhost:8765/?shared_text=タイトル%20https://example.com/
// ============================================================

'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const PORT = Number(process.argv[2]) || 8765;
const ROOT = path.join(__dirname, '..');

// 拡張子 → Content-Type の対応表
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.json': 'application/manifest+json; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.png': 'image/png',
  '.md': 'text/markdown; charset=utf-8'
};

const server = http.createServer((req, res) => {
  // クエリ文字列を除いたパスを取り出す（"/" は index.html にマップ）
  let urlPath = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  if (urlPath === '/') {
    urlPath = '/index.html';
  }

  // パストラバーサル防止: 正規化してルート外へのアクセスを拒否する
  const filePath = path.normalize(path.join(ROOT, urlPath));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found: ' + urlPath);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`開発サーバ起動: http://localhost:${PORT}/`);
  console.log('共有起動の模擬: http://localhost:' + PORT +
    '/?shared_text=' + encodeURIComponent('記事タイトル https://example.com/article'));
});
