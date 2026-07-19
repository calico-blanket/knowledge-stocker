// ============================================================
// サービスワーカー（PWAインストール要件を満たすための最小実装）
//
//   オフライン対応は仕様上不要のため、キャッシュ戦略は持たず
//   すべてネットワークへ素通しする。
// ============================================================

'use strict';

// インストール時: 待機せず即座に有効化する
self.addEventListener('install', function (event) {
  self.skipWaiting();
});

// 有効化時: 既存のタブもすぐ制御下に置く
self.addEventListener('activate', function (event) {
  event.waitUntil(self.clients.claim());
});

// フェッチ: ネットワークへそのまま流す（キャッシュしない）
self.addEventListener('fetch', function (event) {
  event.respondWith(fetch(event.request));
});
