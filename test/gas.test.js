// ============================================================
// gas/Code.gs の自動テスト
//
//   GAS のグローバルサービス（DriveApp / UrlFetchApp 等）を
//   スタブに差し替えた vm コンテキストで Code.gs 本物を実行し、
//   doPost の全経路（正常系・フォールバック・異常系）を検証する。
//
//   このスクリプトは「ナレッジ一覧」スプレッドシートにコンテナバインドされている
//   前提のため、SpreadsheetApp スタブは getActiveSpreadsheet() のみを提供する
//   （openById / create は提供しない。Code.gs がこれらを呼び出せば
//   「関数ではありません」エラーとしてテストが失敗し、回帰を検知できる）。
//
//   実行方法: node --test test/
// ============================================================

'use strict';

const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// ---- GAS サービスのスタブ実装 --------------------------------

/**
 * インメモリの Drive フォルダを作る。
 * GAS の Folder オブジェクトのうち Code.gs が使うメソッドだけ実装する
 * （記事保存はSheetsのみで完結するため、フォルダはDriveショートカット用の
 *   「ナレッジ」ルートフォルダの取得・作成でしか使われない）。
 */
let folderIdCounter = 0;

function createFolderStub(name) {
  const folder = {
    name,
    id: 'folder-' + (++folderIdCounter),
    subFolders: [],
    getId() { return folder.id; },
    isTrashed() { return false; },
    getFoldersByName(target) {
      const hits = folder.subFolders.filter((f) => f.name === target);
      return makeIterator(hits);
    },
    createFolder(target) {
      const child = createFolderStub(target);
      folder.subFolders.push(child);
      return child;
    },
    getUrl() {
      return 'https://drive.google.com/drive/folders/stub-' + encodeURIComponent(name);
    }
  };
  return folder;
}

/**
 * インメモリのスプレッドシート「シート」を作る。
 * 実際のSheetsに合わせ、表示値(getValues)と数式(getFormulas)を別々に保持する。
 * - setFormula: 数式を記録しつつ、=HYPERLINK("url","title") は表示値(title)に反映する
 *   （本物のSheetsが数式を計算表示するのと同じ見え方をNode側で再現するため）
 * - getFormulas: 数式が入ったセルはその数式文字列、それ以外は '' を返す
 *   （fileId抽出は Driveファイル列の HYPERLINK数式から行うため、この再現が必須）
 */
function createSheetStub(maxColumns = 26) {
  const rows = [];       // 表示値（getValues 用）
  const formulas = [];   // 数式（getFormulas 用）。数式でないセルは ''
  // 実スプレッドシートは既定で26列（A〜Z）あり、getRange はこの範囲外を拒否する。
  // ID列（H列=8）の確保処理を検証するため、列数の上限を再現する。
  let columnCount = maxColumns;
  return {
    _rows: rows,
    _formulas: formulas,
    get _columnCount() { return columnCount; },
    appendRow(values) {
      rows.push(values.slice());
      formulas.push(values.map(() => '')); // 追記直後は数式なし
    },
    deleteRow(rowNumber) {
      rows.splice(rowNumber - 1, 1);
      formulas.splice(rowNumber - 1, 1);
    },
    setFrozenRows() {},
    getLastRow() { return rows.length; },
    getMaxColumns() { return columnCount; },
    insertColumnsAfter(afterPosition, howMany) {
      columnCount += howMany;
    },
    getRange(row, col, numRows, numCols) {
      // 実Sheetsと同じく、シートの列数を超える範囲指定は拒否する
      if (col + (numCols || 1) - 1 > columnCount) {
        throw new Error('スタブ: 範囲が列数を超えています (col=' + col + ', max=' + columnCount + ')');
      }
      return {
        setFormula(formula) {
          formulas[row - 1][col - 1] = formula;
          // 表示値も更新（HYPERLINKなら表示テキスト、それ以外は数式文字列のまま）
          const m = formula.match(/HYPERLINK\("((?:[^"]|"")*)","((?:[^"]|"")*)"\)/);
          rows[row - 1][col - 1] = m ? m[2].replace(/""/g, '"') : formula;
        },
        setValue(value) {
          rows[row - 1][col - 1] = value;
          formulas[row - 1][col - 1] = ''; // 値を入れると数式は消える（実挙動と同じ）
        },
        getValue() {
          const line = rows[row - 1];
          return line ? line[col - 1] : undefined;
        },
        // 実Sheetsの getValues は範囲内の未設定セルを '' として返す（undefinedにはならない）。
        // ID列マイグレーションは「空セルかどうか」で判定するため、この再現が必須。
        getValues() {
          const rn = numRows || 1;
          const cn = numCols || 1;
          const out = [];
          for (let r = 0; r < rn; r++) {
            const line = [];
            for (let c = 0; c < cn; c++) {
              const source = rows[row - 1 + r];
              const cell = source ? source[col - 1 + c] : undefined;
              line.push(cell === undefined ? '' : cell);
            }
            out.push(line);
          }
          return out;
        },
        setValues(values) {
          for (let r = 0; r < values.length; r++) {
            const target = row - 1 + r;
            if (!rows[target]) { continue; }
            for (let c = 0; c < values[r].length; c++) {
              rows[target][col - 1 + c] = values[r][c];
              // 値を入れると数式は消える（実挙動と同じ）
              if (formulas[target]) { formulas[target][col - 1 + c] = ''; }
            }
          }
        },
        getFormulas() {
          const rn = numRows || 1;
          const cn = numCols || 1;
          const out = [];
          for (let r = 0; r < rn; r++) {
            const line = [];
            for (let c = 0; c < cn; c++) {
              const fr = formulas[row - 1 + r];
              line.push(fr ? (fr[col - 1 + c] || '') : '');
            }
            out.push(line);
          }
          return out;
        }
      };
    },
    getDataRange() {
      return { getValues: () => rows.map((r) => r.slice()) };
    }
  };
}

/**
 * インメモリの Googleドキュメント「ハンドル」を作る。
 * 記事保存(save)ではもう使わないが、過去に保存されたGoogleドキュメントの
 * 編集(action=update)は引き続きサポートするため、テストの legacy fixture 用に残す。
 */
function createDocStub(id, name) {
  const handle = {
    name,
    content: '',
    getId: () => id,
    getUrl: () => 'https://docs.google.com/document/d/' + id + '/edit',
    getBody() {
      return {
        setText(text) {
          handle.content = text;
          return this;
        },
        getText() { return handle.content; }
      };
    },
    saveAndClose() {}
  };
  return handle;
}

/** フォルダツリーをIDで探す（DriveApp.getFolderById スタブ用） */
function findFolderById(folder, id) {
  if (folder.id === id) { return folder; }
  for (const child of folder.subFolders) {
    const hit = findFolderById(child, id);
    if (hit) { return hit; }
  }
  return null;
}

/** GAS の FolderIterator/FileIterator 相当（hasNext/next だけ） */
function makeIterator(items) {
  let index = 0;
  return {
    hasNext: () => index < items.length,
    next: () => items[index++]
  };
}

/** UrlFetchApp のレスポンス相当を作る */
function makeFetchResponse({ code = 200, body = '', headers = {} }) {
  return {
    getResponseCode: () => code,
    getHeaders: () => headers,
    // charset 指定付きの再読込は、スタブでは body をそのまま返す
    getContentText: (charset) => body
  };
}

// 本番の初期カテゴリ（DEFAULT_CATEGORIES）はユーザーが自分で追加する設計のため空配列。
// テストで category を使う既存ケースの大半はカテゴリの中身自体を検証対象にしていないため、
// ここでテスト用の実在カテゴリを既定でスクリプトプロパティに仕込む（本番の初期値とは独立）。
const TEST_DEFAULT_CATEGORIES = ['PC系', 'DTP系'];

/**
 * Code.gs をスタブ付き vm コンテキストに読み込み、テスト用ハンドルを返す。
 * options.fetchImpl: UrlFetchApp.fetch の差し替え関数
 * options.scriptProperties: スクリプトプロパティの中身
 * options.categories: 初期カテゴリを上書きしたい場合に指定（未指定時は TEST_DEFAULT_CATEGORIES、
 *   scriptProperties.CATEGORIES_JSON を直接指定した場合はそちらを優先）
 */
function loadGasScript(options = {}) {
  const rootFolder = createFolderStub('(root)');
  const fetchCalls = [];
  const docsById = {}; // ドキュメントID -> ハンドル（DocumentApp.create/openById共有）
  let docIdCounter = 0;
  let uuidCounter = 0; // Utilities.getUuid スタブの連番（loadGasScript ごとにリセット）
  const scriptProps = Object.assign({}, options.scriptProperties || {});
  if (!('CATEGORIES_JSON' in scriptProps)) {
    scriptProps.CATEGORIES_JSON = JSON.stringify(options.categories || TEST_DEFAULT_CATEGORIES);
  }

  // コンテナバインド前提: アクティブなスプレッドシート（＝バインド先本体）は1つだけ、
  // 最初は空のシートを1枚持つ状態で存在する。
  const activeSheet = createSheetStub(options.maxColumns);
  const activeSpreadsheet = { getSheets: () => [activeSheet] };

  const context = {
    // --- DriveApp スタブ（Driveショートカット用のルートフォルダ取得のみで使用） ---
    DriveApp: {
      getRootFolder: () => rootFolder,
      getFolderById(id) {
        const hit = findFolderById(rootFolder, id);
        if (!hit || hit === rootFolder) { throw new Error('スタブ: フォルダが見つかりません ' + id); }
        return hit;
      }
    },
    // --- SpreadsheetApp スタブ（コンテナバインド: getActiveSpreadsheetのみ提供） ---
    SpreadsheetApp: {
      getActiveSpreadsheet: () => activeSpreadsheet
    },
    // --- DocumentApp スタブ（過去に保存されたGoogleドキュメントの編集(update)用） ---
    DocumentApp: {
      create(name) {
        const id = 'doc-' + (++docIdCounter);
        const handle = createDocStub(id, name);
        docsById[id] = handle;
        return handle;
      },
      openById(id) {
        const handle = docsById[id];
        if (!handle) {
          throw new Error('スタブ: ドキュメントが見つかりません ' + id);
        }
        return handle;
      }
    },
    // --- UrlFetchApp スタブ ---
    UrlFetchApp: {
      fetch(url, params) {
        fetchCalls.push({ url, params });
        if (options.fetchImpl) {
          return options.fetchImpl(url, params);
        }
        return makeFetchResponse({ body: '<title>デフォルトタイトル</title>' });
      }
    },
    // --- PropertiesService スタブ（getProperty/setProperty とも同じオブジェクトを共有） ---
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => scriptProps[key] || null,
        setProperty: (key, value) => { scriptProps[key] = value; }
      })
    },
    // --- Utilities スタブ（formatDate は固定日時で単純実装） ---
    Utilities: {
      // getUuid は本物と違い連番で再現可能な値を返す（テストで発行順を検証できるようにするため）。
      // 「呼ぶたびに異なる値になる」という本質的な性質は満たしている。
      getUuid() {
        return 'uuid-' + (++uuidCounter);
      },
      formatDate(date, tz, pattern) {
        const pad = (n) => String(n).padStart(2, '0');
        return pattern
          .replace('yyyy', date.getFullYear())
          .replace('MM', pad(date.getMonth() + 1))
          .replace('dd', pad(date.getDate()))
          .replace('HH', pad(date.getHours()))
          .replace('mm', pad(date.getMinutes()))
          .replace('ss', pad(date.getSeconds()));
      }
    },
    // --- ContentService スタブ ---
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(text) {
        const output = {
          _text: text,
          setMimeType: () => output,
          getContent: () => output._text
        };
        return output;
      }
    }
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.gs'), 'utf8');
  vm.runInContext(source, context);

  return { context, rootFolder, fetchCalls, sheet: activeSheet };
}

/** doPost をJSONボディ付きで呼び、レスポンスJSONをパースして返す */
function callDoPost(context, bodyObj) {
  const e = { postData: { contents: JSON.stringify(bodyObj) } };
  const output = context.doPost(e);
  return JSON.parse(output.getContent());
}

/** doGet を ?action=list&category=...&token=...&offset=... 相当のパラメータで呼び、レスポンスJSONを返す */
function callDoGetList(context, category, token, offset) {
  const e = { parameter: { action: 'list', category: category, token: token, offset: offset } };
  const output = context.doGet(e);
  return JSON.parse(output.getContent());
}

/** doGet を ?action=categories&token=... で呼び、レスポンスJSONを返す */
function callDoGetCategories(context, token) {
  const output = context.doGet({ parameter: { action: 'categories', token: token } });
  return JSON.parse(output.getContent());
}

/**
 * 「移行前（コンテナバインド化・タグ機能追加より前）に保存された記事」を
 * インデックスシートとGoogleドキュメントの両方に直接作り込む。
 * 現在の save フローはもうGoogleドキュメントを作らないため、
 * action=update（既存ドキュメントの編集）をテストするにはこの関数で
 * legacy な状態を再現する。
 */
function seedLegacyArticle(context, { savedAt, category, title, url, memo }) {
  const sheet = context.getOrCreateIndexSheet_();
  const doc = context.DocumentApp.create(title);
  doc.getBody().setText(context.buildMarkdown_(title, url, savedAt, category, memo || '', '', ''));

  sheet.appendRow([savedAt, category, title, url, memo || '', '', '']);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 3).setFormula('=HYPERLINK("' + url + '","' + title + '")');
  sheet.getRange(lastRow, 6).setFormula('=HYPERLINK("' + doc.getUrl() + '","開く")');

  return { fileId: doc.getId(), sheet, lastRow };
}

// ---- 正常系（保存はSheetsインデックスのみ・Googleドキュメントは作らない） ----

test('正常系: タイトル取得 → Sheetsインデックスに1行追記される（Googleドキュメントは作らない）', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<html><head><title>テスト記事のタイトル</title></head></html>' })
  });

  const result = callDoPost(context, { url: 'https://example.com/article', category: 'PC系' });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.title, 'テスト記事のタイトル');
  assert.strictEqual(result.category, 'PC系');
  assert.deepStrictEqual(Array.from(result.tags), []);

  assert.strictEqual(sheet._rows.length, 2, 'ヘッダ行 + データ1行');
  assert.deepStrictEqual(
    Array.from(sheet._rows[0]),
    ['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']
  );
  assert.strictEqual(sheet._rows[1][1], 'PC系');
  assert.strictEqual(sheet._rows[1][2], 'テスト記事のタイトル'); // HYPERLINKの表示値
  assert.strictEqual(sheet._rows[1][3], 'https://example.com/article');
  assert.strictEqual(sheet._rows[1][5], '', 'Driveファイル列は新規保存では常に空');
});

test('正常系: メモ付きで保存するとメモ列に反映される', () => {
  const { context, sheet } = loadGasScript();
  const result = callDoPost(context, {
    url: 'https://example.com/', category: 'PC系', memo: 'あとで読む'
  });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][4], 'あとで読む');
});

test('正常系: 2回目以降の保存でもヘッダー行は重複せず、行が追記されていく', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/1', category: 'DTP系' });
  callDoPost(context, { url: 'https://example.com/2', category: 'DTP系' });

  assert.strictEqual(sheet._rows.length, 3, 'ヘッダ + データ2行');
  assert.deepStrictEqual(Array.from(sheet._rows[0]), ['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
});

// ---- タイトル取得のフォールバック -----------------------------

test('フォールバック: fetch が例外を投げたら URL がタイトルになる', () => {
  const { context } = loadGasScript({
    fetchImpl: () => { throw new Error('DNS解決失敗'); }
  });
  const result = callDoPost(context, { url: 'https://unreachable.example/', category: 'PC系' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.title, 'https://unreachable.example/');
});

test('フォールバック: HTTP 404 なら URL がタイトルになる', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ code: 404, body: 'Not Found' })
  });
  const result = callDoPost(context, { url: 'https://example.com/gone', category: 'PC系' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.title, 'https://example.com/gone');
});

test('フォールバック: <title> が無い HTML なら og:title を使う', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({
      body: '<meta property="og:title" content="OGタイトルだけのページ"><p>本文</p>'
    })
  });
  const result = callDoPost(context, { url: 'https://example.com/og', category: 'PC系' });
  assert.strictEqual(result.title, 'OGタイトルだけのページ');
});

test('フォールバック: <title> も og:title も無ければ URL を使う', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<p>タイトルなし</p>' })
  });
  const result = callDoPost(context, { url: 'https://example.com/notitle', category: 'PC系' });
  assert.strictEqual(result.title, 'https://example.com/notitle');
});

// ---- タイトル抽出の純粋関数 -----------------------------------

test('extractTitle_: HTML実体参照がデコードされ、改行・連続空白が畳まれる', () => {
  const { context } = loadGasScript();
  const title = context.extractTitle_(
    '<title>\n  A &amp; B &lt;C&gt; &quot;D&quot; &#39;E&#x27;   F\n</title>'
  );
  assert.strictEqual(title, 'A & B <C> "D" \'E\' F');
});

test('extractTitle_: 属性付き title タグでも抽出できる', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.extractTitle_('<title data-rh="true">属性付き</title>'), '属性付き');
});

// ---- 異常系（入力バリデーション） ------------------------------

test('異常系: URL 無しはエラー', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { category: 'PC系' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /URLが指定されていません/);
});

test('異常系: javascript: スキームは拒否される', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { url: 'javascript:alert(1)', category: 'PC系' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /URLの形式が不正です/);
});

test('異常系: 許可リストにないカテゴリは拒否される', () => {
  const { context, sheet } = loadGasScript();
  const result = callDoPost(context, { url: 'https://example.com/', category: '../etc' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /不明なカテゴリです/);
  assert.strictEqual(sheet._rows.length, 0, '行は追記されない');
});

test('異常系: ボディが JSON でない場合はエラー', () => {
  const { context } = loadGasScript();
  const output = context.doPost({ postData: { contents: 'not-json' } });
  const result = JSON.parse(output.getContent());
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /JSONが不正/);
});

test('異常系: ボディ無し（e が空）でもエラーJSONを返して落ちない', () => {
  const { context } = loadGasScript();
  const output = context.doPost(undefined);
  const result = JSON.parse(output.getContent());
  assert.strictEqual(result.ok, false);
});

// ---- 合言葉（SHARED_TOKEN） -----------------------------------

test('token: SHARED_TOKEN 設定時、一致しないと拒否される', () => {
  const { context, sheet } = loadGasScript({
    scriptProperties: { SHARED_TOKEN: 'himitsu' }
  });
  const ng = callDoPost(context, { url: 'https://example.com/', category: 'PC系', token: 'wrong' });
  assert.strictEqual(ng.ok, false);
  assert.match(ng.error, /合言葉が一致しません/);
  assert.strictEqual(sheet._rows.length, 0, '保存されない');

  const ok = callDoPost(context, { url: 'https://example.com/', category: 'PC系', token: 'himitsu' });
  assert.strictEqual(ok.ok, true);
});

test('token: SHARED_TOKEN 未設定なら token 無しでも通る', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { url: 'https://example.com/', category: 'PC系' });
  assert.strictEqual(result.ok, true);
});

// ---- 全カテゴリの網羅確認 -------------------------------------

test('カテゴリ: 設定済みの全カテゴリが受理され、Sheetsに1行ずつ追記される', () => {
  const expected = ['旅行', '料理', 'ガジェット', '読書'];
  const { context, sheet } = loadGasScript({ categories: expected });
  for (const cat of expected) {
    const result = callDoPost(context, { url: 'https://example.com/', category: cat });
    assert.strictEqual(result.ok, true, cat + ' が受理される');
  }
  assert.strictEqual(sheet._rows.length, 1 + expected.length, 'ヘッダ + ' + expected.length + '行');
});

// ---- 一覧API -----------------------------------------------------

test('一覧API: 指定カテゴリの保存済み記事のみを新しい順に返す', () => {
  const { context } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/pc1', category: 'PC系', memo: '1件目' });
  callDoPost(context, { url: 'https://example.com/dtp1', category: 'DTP系' });
  callDoPost(context, { url: 'https://example.com/pc2', category: 'PC系', memo: '2件目' });

  const result = callDoGetList(context, 'PC系');
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.items.length, 2, 'PC系のみ2件');
  // 新しい順（後に保存したpc2が先頭）
  assert.strictEqual(result.items[0].url, 'https://example.com/pc2');
  assert.strictEqual(result.items[0].memo, '2件目');
  assert.strictEqual(result.items[1].url, 'https://example.com/pc1');
});

test('一覧API: タグ無し記事は items[].tags が空配列', () => {
  const { context } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/no-tag', category: 'PC系' });

  const result = callDoGetList(context, 'PC系');
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.items[0].tags), []);
});

test('一覧API: タグ1件の記事は items[].tags にそのタグが1件入る', () => {
  const { context } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: 'サンプルタグ' });
  callDoPost(context, { url: 'https://example.com/one-tag', category: 'PC系', tags: ['サンプルタグ'] });

  const result = callDoGetList(context, 'PC系');
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.items[0].tags), ['サンプルタグ']);
});

test('一覧API: タグ複数件の記事は items[].tags に保存順で全件入る', () => {
  const { context } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: 'タグ1' });
  callDoPost(context, { action: 'addTag', name: 'タグ2' });
  callDoPost(context, { action: 'addTag', name: 'タグ3' });
  callDoPost(context, {
    url: 'https://example.com/multi-tag', category: 'PC系', tags: ['タグ1', 'タグ2', 'タグ3']
  });

  const result = callDoGetList(context, 'PC系');
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.items[0].tags), ['タグ1', 'タグ2', 'タグ3']);
});

test('一覧API: まだ何も保存されていないカテゴリは空配列を返す（エラーにしない）', () => {
  const { context } = loadGasScript();
  const result = callDoGetList(context, 'DTP系');
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(result.items, []);
});

test('一覧API: 不明なカテゴリはエラーを返す', () => {
  const { context } = loadGasScript();
  const result = callDoGetList(context, '../etc');
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /不明なカテゴリです/);
});

test('一覧API: category未指定はエラーを返す', () => {
  const { context } = loadGasScript();
  const result = callDoGetList(context, undefined);
  assert.strictEqual(result.ok, false);
});

test('doGet: action未指定は稼働確認メッセージを返す（一覧処理に入らない）', () => {
  const { context } = loadGasScript();
  const output = context.doGet({ parameter: {} });
  const result = JSON.parse(output.getContent());
  assert.strictEqual(result.ok, true);
  assert.match(result.message, /稼働中/);
});

test('doGet: eが空でも落ちない', () => {
  const { context } = loadGasScript();
  const output = context.doGet(undefined);
  const result = JSON.parse(output.getContent());
  assert.strictEqual(result.ok, true);
});

test('escapeFormulaString_: ダブルクォートが二重化される（数式インジェクション対策）', () => {
  const { context } = loadGasScript();
  assert.strictEqual(
    context.escapeFormulaString_('タイトルに"引用符"あり'),
    'タイトルに""引用符""あり'
  );
});

// ---- カテゴリ管理（動的CRUD） -----------------------------------

test('categories API: 設定済みカテゴリ・空のタグ一覧が返る', () => {
  const { context } = loadGasScript();
  const result = callDoGetCategories(context);
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.categories), TEST_DEFAULT_CATEGORIES);
  assert.deepStrictEqual(Array.from(result.tags), [], 'タグは最初は空（設定画面から追加していく運用）');
  assert.match(result.rootFolderUrl, /^https:\/\/drive\.google\.com\/drive\/folders\//);
});

test('categories API: 2回呼んでも同じ一覧が安定して返る（毎回初期化されない）', () => {
  const { context } = loadGasScript();
  const first = callDoGetCategories(context);
  const second = callDoGetCategories(context);
  assert.deepStrictEqual(Array.from(first.categories), Array.from(second.categories));
});

test('addCategory: 新しいカテゴリを追加すると一覧に反映され、以後 save でも使える', () => {
  const { context } = loadGasScript();
  const added = callDoPost(context, { action: 'addCategory', name: '写真' });
  assert.strictEqual(added.ok, true);
  assert.ok(added.categories.indexOf('写真') !== -1);

  const saved = callDoPost(context, { url: 'https://example.com/photo', category: '写真' });
  assert.strictEqual(saved.ok, true, '追加した直後のカテゴリで保存できる');
});

test('addCategory: 空名は拒否される', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'addCategory', name: '   ' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /カテゴリ名が指定されていません/);
});

test('addCategory: 同名カテゴリの重複追加は拒否される', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'addCategory', name: 'PC系' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /同名のカテゴリが既にあります/);
});

test('addCategory: 名前に / や \\ が含まれる場合は・に置換される（Driveフォルダ名対策）', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'addCategory', name: 'A/B\\C' });
  assert.strictEqual(result.ok, true);
  assert.ok(result.categories.indexOf('A・B・C') !== -1);
});

test('removeCategory: 削除すると一覧から消えるが、既存のSheets行は残る', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/old', category: 'PC系', memo: '削除前に保存' });

  const removed = callDoPost(context, { action: 'removeCategory', name: 'PC系' });
  assert.strictEqual(removed.ok, true);
  assert.strictEqual(removed.categories.indexOf('PC系'), -1);

  // 新規保存はもうできない（選択肢から外れている）
  const saveAfterRemove = callDoPost(context, { url: 'https://example.com/new', category: 'PC系' });
  assert.strictEqual(saveAfterRemove.ok, false);

  // 削除前に保存した行は残っている（非破壊）
  assert.strictEqual(sheet._rows.length, 2, 'ヘッダ + 削除前の1行は残る');
});

test('removeCategory: 存在しないカテゴリの削除はエラー', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'removeCategory', name: '存在しないカテゴリ' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /存在しないカテゴリです/);
});

test('doPost: 不明なactionはエラーになる', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'destroyEverything' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /不明なactionです/);
});

test('doPost: addCategory/removeCategoryもSHARED_TOKEN検証の対象になる', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const result = callDoPost(context, { action: 'addCategory', name: '写真', token: 'wrong' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
});

test('DEFAULT_CATEGORIES: 本番の初期カテゴリは空配列（ユーザーが自分で追加する設計）', () => {
  const { context } = loadGasScript();
  assert.deepStrictEqual(Array.from(context.DEFAULT_CATEGORIES), []);
});

test('categories API: 本番の初期カテゴリ（スクリプトプロパティ未設定）は空配列で返る', () => {
  const { context } = loadGasScript({ scriptProperties: {} , categories: [] });
  const result = callDoGetCategories(context);
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.categories), []);
});

// ---- タグ管理（動的CRUD。カテゴリと同じ方式だが初期値は空配列） ----------

test('getTags_: 初回アクセスでは空配列を返す（カテゴリと違い初期候補は無い）', () => {
  const { context } = loadGasScript();
  assert.deepStrictEqual(Array.from(context.getTags_()), []);
});

test('addTag: 新しいタグを追加すると一覧に反映され、以後 save でも使える', () => {
  const { context, sheet } = loadGasScript();
  const added = callDoPost(context, { action: 'addTag', name: '要対応' });
  assert.strictEqual(added.ok, true);
  assert.deepStrictEqual(Array.from(added.tags), ['要対応']);

  const saved = callDoPost(context, {
    url: 'https://example.com/x', category: 'PC系', tags: ['要対応']
  });
  assert.strictEqual(saved.ok, true, '追加した直後のタグで保存できる');
  assert.strictEqual(sheet._rows[1][6], '要対応', 'タグ列に反映される');
});

test('addTag: 空名は拒否される', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'addTag', name: '   ' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /タグ名が指定されていません/);
});

test('addTag: 同名タグの重複追加は拒否される', () => {
  const { context } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: '要対応' });
  const result = callDoPost(context, { action: 'addTag', name: '要対応' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /同名のタグが既にあります/);
});

test('addTag: 名前にカンマが含まれる場合は読点に置換される（タグ列のカンマ区切り対策）', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'addTag', name: 'A,B' });
  assert.strictEqual(result.ok, true);
  assert.ok(result.tags.indexOf('A、B') !== -1);
});

test('removeTag: 削除すると一覧から消えるが、既存のSheets行のタグ文字列は変わらない', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: '要対応' });
  callDoPost(context, { url: 'https://example.com/x', category: 'PC系', tags: ['要対応'] });

  const removed = callDoPost(context, { action: 'removeTag', name: '要対応' });
  assert.strictEqual(removed.ok, true);
  assert.strictEqual(removed.tags.indexOf('要対応'), -1);

  // 新規保存でこのタグはもう使えない
  const saveAfterRemove = callDoPost(context, {
    url: 'https://example.com/y', category: 'PC系', tags: ['要対応']
  });
  assert.strictEqual(saveAfterRemove.ok, false);
  assert.match(saveAfterRemove.error, /不明なタグです/);

  // 削除前に保存した行のタグ列はそのまま
  assert.strictEqual(sheet._rows[1][6], '要対応');
});

test('removeTag: 存在しないタグの削除はエラー', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'removeTag', name: '存在しないタグ' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /存在しないタグです/);
});

test('reorderTags: 並び替えた順序が保存され、以後の categories API に反映される', () => {
  const { context } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: 'A' });
  callDoPost(context, { action: 'addTag', name: 'B' });
  callDoPost(context, { action: 'addTag', name: 'C' });

  const result = callDoPost(context, { action: 'reorderTags', tags: ['C', 'A', 'B'] });
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(result.tags, ['C', 'A', 'B']);
  assert.deepStrictEqual(callDoGetCategories(context).tags, ['C', 'A', 'B']);
});

test('reorderTags: 件数や中身が一致しない並び替えは拒否される', () => {
  const { context } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: 'A' });
  callDoPost(context, { action: 'addTag', name: 'B' });

  const missingOne = callDoPost(context, { action: 'reorderTags', tags: ['A'] });
  assert.strictEqual(missingOne.ok, false);

  const renamed = callDoPost(context, { action: 'reorderTags', tags: ['A', '違う名前'] });
  assert.strictEqual(renamed.ok, false);

  assert.deepStrictEqual(callDoGetCategories(context).tags, ['A', 'B']);
});

test('reorderTags: tags が配列でない・無い場合はエラー', () => {
  const { context } = loadGasScript();
  const r1 = callDoPost(context, { action: 'reorderTags' });
  assert.strictEqual(r1.ok, false);
  const r2 = callDoPost(context, { action: 'reorderTags', tags: 'A' });
  assert.strictEqual(r2.ok, false);
});

test('doPost: addTag/removeTag/reorderTagsもSHARED_TOKEN検証の対象になる', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const result = callDoPost(context, { action: 'addTag', name: 'A', token: 'wrong' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
});

test('save: 複数タグを選択して保存すると、Sheetsのタグ列に ", " 区切りでまとめられる', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { action: 'addTag', name: 'タグA' });
  callDoPost(context, { action: 'addTag', name: 'タグB' });

  const result = callDoPost(context, {
    url: 'https://example.com/multi', category: 'PC系', tags: ['タグA', 'タグB']
  });
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.tags), ['タグA', 'タグB']);
  assert.strictEqual(sheet._rows[1][6], 'タグA, タグB');
});

test('save: タグ未指定・空配列でも保存できる（任意項目）', () => {
  const { context, sheet } = loadGasScript();
  const result = callDoPost(context, { url: 'https://example.com/notag', category: 'PC系' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][6], '');
});

test('save: 現在のタグ一覧に無いタグを指定するとエラーになる', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, {
    url: 'https://example.com/x', category: 'PC系', tags: ['存在しないタグ']
  });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /不明なタグです/);
});

test('normalizeTagsInput_: 配列以外は空配列に、前後空白・空文字は除去される', () => {
  const { context } = loadGasScript();
  // vm(別レルム)の配列と比較するため、prototypeを問わない Array.from で正規化してから比較する
  assert.deepStrictEqual(Array.from(context.normalizeTagsInput_(undefined)), []);
  assert.deepStrictEqual(Array.from(context.normalizeTagsInput_('タグA')), []);
  assert.deepStrictEqual(
    Array.from(context.normalizeTagsInput_(['  タグA  ', '', 'タグB'])), ['タグA', 'タグB']
  );
});

// ---- コンテナバインド化（getOrCreateIndexSheet_） -----------------------

test('コンテナバインド: getActiveSpreadsheetの最初のシートを使い、空なら見出し行を書き込む', () => {
  const { context, sheet } = loadGasScript();
  const returned = context.getOrCreateIndexSheet_();
  assert.strictEqual(returned, sheet, 'アクティブなスプレッドシートの1枚目のシートを返す');
  assert.deepStrictEqual(
    Array.from(sheet._rows[0]),
    ['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']
  );
});

test('コンテナバインド: 列追加前から運用しているシート（タグ列ヘッダー無し）はヘッダーだけ補完する', () => {
  const { context, sheet } = loadGasScript();
  // 移行前の6列ヘッダーを直接書き込んでおく（タグ列は存在しない状態を再現）
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル']);
  sheet.appendRow(['2026-07-01 10:00', 'PC系', '既存記事', 'https://example.com/old', '', '']);

  context.getOrCreateIndexSheet_();

  assert.strictEqual(sheet._rows[0][6], 'タグ', 'G1にヘッダーが補完される');
  assert.strictEqual(sheet._rows.length, 2, '既存データ行は変更されない');
  assert.strictEqual(sheet._rows[1][2], '既存記事', '既存データ行は変更されない');
});

// ---- URL解決（短縮/リダイレクトリンク対策） -----------------------

test('resolveFinalUrl_: リダイレクトが無ければ元のURLをそのまま返す', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ code: 200, body: 'ok' })
  });
  assert.strictEqual(context.resolveFinalUrl_('https://example.com/'), 'https://example.com/');
});

test('resolveFinalUrl_: 302リダイレクトを1回辿って実URLへ解決する', () => {
  const { context } = loadGasScript({
    fetchImpl: (url) => {
      if (url === 'https://short.example/abc') {
        return makeFetchResponse({ code: 302, headers: { Location: 'https://real.example/article' } });
      }
      return makeFetchResponse({ code: 200 });
    }
  });
  assert.strictEqual(context.resolveFinalUrl_('https://short.example/abc'), 'https://real.example/article');
});

test('resolveFinalUrl_: リダイレクトが上限回数を超えたら最後に辿り着いたURLで打ち切る', () => {
  const { context } = loadGasScript({
    fetchImpl: (url) => {
      const n = Number(url.split('/').pop());
      return makeFetchResponse({ code: 302, headers: { Location: 'https://loop.example/' + (n + 1) } });
    }
  });
  // MAX_REDIRECT_HOPS=5 なので、0→1→2→3→4→5 と5回転送を辿った時点で打ち切られる
  assert.strictEqual(context.resolveFinalUrl_('https://loop.example/0'), 'https://loop.example/5');
});

test('resolveFinalUrl_: Locationヘッダが無い3xxはその時点のURLで打ち切る', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ code: 301, headers: {} })
  });
  assert.strictEqual(context.resolveFinalUrl_('https://example.com/'), 'https://example.com/');
});

test('resolveFinalUrl_: 通信エラー時はその時点のURLを返す（例外を投げない）', () => {
  const { context } = loadGasScript({
    fetchImpl: () => { throw new Error('timeout'); }
  });
  assert.strictEqual(context.resolveFinalUrl_('https://example.com/'), 'https://example.com/');
});

test('resolveRelativeUrl_: 絶対URLのLocationはそのまま返す', () => {
  const { context } = loadGasScript();
  assert.strictEqual(
    context.resolveRelativeUrl_('https://a.example/x', 'https://b.example/y'),
    'https://b.example/y'
  );
});

test('resolveRelativeUrl_: /始まりの相対パスはoriginと結合する', () => {
  const { context } = loadGasScript();
  assert.strictEqual(
    context.resolveRelativeUrl_('https://a.example/x/y', '/z'),
    'https://a.example/z'
  );
});

test('doPost: 短縮/リダイレクトURLは実URLに解決されてからSheetsに保存される', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: (url) => {
      if (url === 'https://short.example/abc') {
        return makeFetchResponse({ code: 302, headers: { Location: 'https://real.example/article' } });
      }
      return makeFetchResponse({ body: '<title>実記事タイトル</title><p>本文だよ</p>' });
    }
  });

  const result = callDoPost(context, { url: 'https://short.example/abc', category: 'PC系' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.title, '実記事タイトル');
  assert.strictEqual(sheet._rows[1][3], 'https://real.example/article', 'Sheets側のURL列は解決後のURL');
});

// ---- メモ重複排除（タイトルと同一のメモを捨てる） -------------------

test('isDuplicateMemo_: メモとタイトルが完全一致なら重複と判定される', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.isDuplicateMemo_('同じ文字列', '同じ文字列'), true);
});

test('isDuplicateMemo_: 前後の空白差は無視して重複判定する', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.isDuplicateMemo_('  同じ文字列  ', '同じ文字列'), true);
});

test('isDuplicateMemo_: 内容が異なれば重複ではない', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.isDuplicateMemo_('あとで読む', 'タイトル'), false);
});

test('isDuplicateMemo_: 空メモは重複ではない', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.isDuplicateMemo_('', 'タイトル'), false);
});

test('doPost: メモがタイトルと同一なら、Sheetsのメモ列は空になる（重複防止）', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<title>同じ内容</title>' })
  });
  const result = callDoPost(context, { url: 'https://example.com/dup', category: 'PC系', memo: '同じ内容' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][4], '');
});

test('doPost: メモがタイトルと異なれば、Sheetsのメモ列にそのまま入る', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<title>タイトル</title>' })
  });
  callDoPost(context, { url: 'https://example.com/diff', category: 'PC系', memo: '別のコメント' });
  assert.strictEqual(sheet._rows[1][4], '別のコメント');
});

// ---- GET系APIの合言葉検証 -----------------------------------------
// 記事一覧はタイトル・URL・メモといった個人の閲覧記録に近い情報を含むため、
// WebアプリのURLを知られただけでは読めないよう、GETもPOSTと同じ合言葉で保護する。

test('GET保護: SHARED_TOKEN設定時、tokenなしのlistは拒否される', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const result = callDoGetList(context, 'PC系', undefined);
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
});

test('GET保護: SHARED_TOKEN設定時、token不一致のcategoriesは拒否される', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const result = callDoGetCategories(context, 'wrong');
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
});

test('GET保護: 正しいtokenならlist/categoriesとも通る', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系', token: 'himitsu' });

  const list = callDoGetList(context, 'PC系', 'himitsu');
  assert.strictEqual(list.ok, true);
  assert.strictEqual(list.items.length, 1);

  const categories = callDoGetCategories(context, 'himitsu');
  assert.strictEqual(categories.ok, true);
});

test('GET保護: SHARED_TOKEN未設定ならtokenなしでも通る（従来挙動の維持）', () => {
  const { context } = loadGasScript();
  const result = callDoGetCategories(context, undefined);
  assert.strictEqual(result.ok, true);
});

test('GET保護: tokenなしでも稼働確認メッセージ(actionなし)は返る', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const output = context.doGet({ parameter: {} });
  const result = JSON.parse(output.getContent());
  assert.strictEqual(result.ok, true, '稼働確認は情報を含まないため合言葉不要');
});

// ---- Sheets数式インジェクション対策 ---------------------------------

test('sanitizeCellText_: 先頭が = の文字列にはアポストロフィが付く', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.sanitizeCellText_('=IMPORTXML("http://evil/","//a")'), "'=IMPORTXML(\"http://evil/\",\"//a\")");
  assert.strictEqual(context.sanitizeCellText_('+1+1'), "'+1+1");
});

test('sanitizeCellText_: 通常の文字列はそのまま', () => {
  const { context } = loadGasScript();
  assert.strictEqual(context.sanitizeCellText_('普通のタイトル'), '普通のタイトル');
  assert.strictEqual(context.sanitizeCellText_(''), '');
});

test('doPost: メモが数式で始まる場合、Sheetsにはアポストロフィ付きで書き込まれる', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<title>数式入りメモの記事</title>' })
  });
  callDoPost(context, { url: 'https://example.com/f', category: 'PC系', memo: '=1+1' });
  assert.strictEqual(sheet._rows[1][4], "'=1+1", 'メモ列は数式として解釈されない形で格納される');
});

// ---- コード品質の回帰チェック（レビュー指摘🔴3） ---------------------

test('Code.gs: checkToken_ の定義がちょうど1つである（重複定義の再発防止）', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.gs'), 'utf8');
  const definitions = source.match(/function checkToken_\(/g) || [];
  assert.strictEqual(definitions.length, 1);
});

// ---- カテゴリの並び替え(action=reorderCategories) --------------------

test('reorderCategories: 並び替えた順序が保存され、以後の categories API に反映される', () => {
  const { context } = loadGasScript();
  const original = callDoGetCategories(context).categories;
  const reordered = original.slice().reverse();

  const result = callDoPost(context, { action: 'reorderCategories', categories: reordered });
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(result.categories, reordered);

  // 再取得しても新しい順序が保持されている（スクリプトプロパティに永続化）
  assert.deepStrictEqual(callDoGetCategories(context).categories, reordered);
});

test('reorderCategories: 件数が一致しない（勝手な削除・追加の混入）は拒否される', () => {
  const { context } = loadGasScript();
  const original = callDoGetCategories(context).categories;

  const missingOne = original.slice(1); // 1件欠け
  const r1 = callDoPost(context, { action: 'reorderCategories', categories: missingOne });
  assert.strictEqual(r1.ok, false);
  assert.match(r1.error, /一致しません/);

  const extraOne = original.concat(['勝手に追加']); // 1件過剰
  const r2 = callDoPost(context, { action: 'reorderCategories', categories: extraOne });
  assert.strictEqual(r2.ok, false);

  // どちらの失敗でも元の並びが壊れていないこと
  assert.deepStrictEqual(callDoGetCategories(context).categories, original);
});

test('reorderCategories: 同数でも要素の中身が違う（改名の混入）は拒否される', () => {
  const { context } = loadGasScript();
  const original = callDoGetCategories(context).categories;
  const renamed = original.slice();
  renamed[0] = '存在しない名前へ改名';

  const result = callDoPost(context, { action: 'reorderCategories', categories: renamed });
  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(callDoGetCategories(context).categories, original);
});

test('reorderCategories: categories が配列でない・無い場合はエラー', () => {
  const { context } = loadGasScript();
  const r1 = callDoPost(context, { action: 'reorderCategories' });
  assert.strictEqual(r1.ok, false);
  const r2 = callDoPost(context, { action: 'reorderCategories', categories: 'PC系' });
  assert.strictEqual(r2.ok, false);
});

test('reorderCategories: SHARED_TOKEN検証の対象になる（合言葉なしは拒否）', () => {
  const { context } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'himitsu' } });
  const result = callDoPost(context, {
    action: 'reorderCategories',
    categories: ['a'] // tokenチェックが先に走るため中身は届かない
  });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
});

test('reorderCategories: 並び替え後に追加したカテゴリは従来どおり末尾に付く', () => {
  const { context } = loadGasScript();
  const original = callDoGetCategories(context).categories;
  const reordered = original.slice().reverse();
  callDoPost(context, { action: 'reorderCategories', categories: reordered });

  const result = callDoPost(context, { action: 'addCategory', name: '新しいカテゴリ' });
  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(result.categories, reordered.concat(['新しいカテゴリ']));
});

// ---- 一覧APIのページング(50件区切り + offset) ------------------------

test('一覧API: 50件を超える保存がある場合、先頭ページは新しい順50件 + hasMore:true になる', () => {
  const { context } = loadGasScript({
    fetchImpl: (url) => makeFetchResponse({ body: '<title>記事' + url.split('/').pop() + '</title>' })
  });
  // 55件保存する（記事0が最古、記事54が最新）
  for (let i = 0; i < 55; i++) {
    const r = callDoPost(context, { url: 'https://example.com/' + i, category: 'PC系' });
    assert.strictEqual(r.ok, true, i + '件目の保存が成功すること');
  }

  const page1 = callDoGetList(context, 'PC系');
  assert.strictEqual(page1.ok, true);
  assert.strictEqual(page1.items.length, 50, '先頭ページは50件で区切られる');
  assert.strictEqual(page1.hasMore, true, '続きがあることを示すフラグが立つ');
  assert.strictEqual(page1.items[0].title, '記事54', '最新の保存が先頭に来る');
  assert.strictEqual(page1.items[49].title, '記事5', '50件目は新しい順で50番目');
});

test('一覧API: offset指定で続きのページが取得でき、最後のページは hasMore:false になる', () => {
  const { context } = loadGasScript({
    fetchImpl: (url) => makeFetchResponse({ body: '<title>記事' + url.split('/').pop() + '</title>' })
  });
  for (let i = 0; i < 55; i++) {
    callDoPost(context, { url: 'https://example.com/' + i, category: 'PC系' });
  }

  const page2 = callDoGetList(context, 'PC系', undefined, '50');
  assert.strictEqual(page2.ok, true);
  assert.strictEqual(page2.items.length, 5, '2ページ目は残りの5件');
  assert.strictEqual(page2.hasMore, false, '最後のページではフラグが下りる');
  assert.strictEqual(page2.items[0].title, '記事4');
  assert.strictEqual(page2.items[4].title, '記事0', '最古の保存が末尾に来る');
});

test('一覧API: offsetが不正な値（負数・文字列）の場合は先頭ページとして扱う', () => {
  const { context } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<title>1件だけ</title>' })
  });
  callDoPost(context, { url: 'https://example.com/only', category: 'PC系' });

  for (const bad of ['-5', 'abc', '']) {
    const result = callDoGetList(context, 'PC系', undefined, bad);
    assert.strictEqual(result.ok, true, 'offset=' + JSON.stringify(bad) + ' でもエラーにしない');
    assert.strictEqual(result.items.length, 1);
    assert.strictEqual(result.hasMore, false);
  }
});

// ---- ルートフォルダIDのキャッシュ（カテゴリ・一覧APIの高速化） --------

test('ルートフォルダ: 一度アクセスするとIDがスクリプトプロパティに記憶される', () => {
  const { context } = loadGasScript();
  callDoGetCategories(context);
  const props = context.PropertiesService.getScriptProperties();
  assert.ok(props.getProperty('ROOT_FOLDER_ID'), 'ROOT_FOLDER_IDが保存されること');
});

test('ルートフォルダ: 記憶されたIDのフォルダが消えていても名前検索にフォールバックして復旧する', () => {
  const { context, rootFolder } = loadGasScript({
    scriptProperties: { ROOT_FOLDER_ID: 'folder-存在しないID' }
  });
  const result = callDoGetCategories(context);
  assert.strictEqual(result.ok, true, '古いIDが無効でもエラーにならない');
  assert.strictEqual(rootFolder.subFolders.length, 1, 'ナレッジフォルダが作られる');
  const props = context.PropertiesService.getScriptProperties();
  assert.strictEqual(
    props.getProperty('ROOT_FOLDER_ID'), rootFolder.subFolders[0].getId(),
    '正しいIDに更新されること'
  );
});

test('ルートフォルダ: IDキャッシュ利用時もフォルダが二重に作られない', () => {
  const { context, rootFolder } = loadGasScript({
    fetchImpl: () => makeFetchResponse({ body: '<title>a</title>' })
  });
  callDoPost(context, { url: 'https://example.com/1', category: 'PC系' });
  callDoGetCategories(context);
  callDoGetList(context, 'PC系');
  callDoPost(context, { url: 'https://example.com/2', category: 'PC系' });

  const knowledgeFolders = rootFolder.subFolders.filter((f) => f.name === 'ナレッジ');
  assert.strictEqual(knowledgeFolders.length, 1, '「ナレッジ」フォルダは1つだけ');
});

// ---- 保存済み記事（過去のGoogleドキュメント）の編集（action=update） --------
// 現在の save はもうGoogleドキュメントを作らないため、editテストは
// seedLegacyArticle() で「移行前に保存された記事」を直接作り込んでから検証する。

test('fileId抽出: Driveファイル列のHYPERLINK数式からfileIdを取り出せる', () => {
  const { context } = loadGasScript();
  assert.strictEqual(
    context.extractFileIdFromFormula_('=HYPERLINK("https://docs.google.com/document/d/ABC_12-3/edit","開く")'),
    'ABC_12-3'
  );
  assert.strictEqual(context.extractFileIdFromFormula_(''), '', '数式なしは空文字');
  assert.strictEqual(context.extractFileIdFromFormula_(null), '', 'nullも空文字');
  assert.strictEqual(
    context.extractFileIdFromFormula_('=HYPERLINK("https://example.com/","開く")'),
    '', 'ドキュメントURL形式でない数式は空文字'
  );
});

test('Doc本文抽出: 共有時のURL行と本文（自動抽出・参考）節を取り出せる（無ければ空文字）', () => {
  const { context } = loadGasScript();
  const docText = [
    '# 旧タイトル',
    '',
    '- URL: https://old.example.com/a',
    '- 共有時のURL（短縮/リダイレクト元）: https://share.google/xyz',
    '- 保存日時: 2026-07-01 10:00',
    '- カテゴリ: PC系',
    '',
    '## 本文（自動抽出・参考）',
    '',
    'これは抽出された本文です。',
    ''
  ].join('\n');

  assert.strictEqual(context.extractDocOriginalUrl_(docText), 'https://share.google/xyz');
  assert.strictEqual(context.extractDocBodyText_(docText), 'これは抽出された本文です。');
  assert.strictEqual(context.extractDocOriginalUrl_('# タイトルのみ'), '', '共有時URL行が無ければ空文字');
  assert.strictEqual(context.extractDocBodyText_('# タイトルのみ'), '', '本文節が無ければ空文字');
});

test('Doc再構築: タイトル・URL・メモを差し替えつつ、共有時URLと自動抽出本文は引き継ぐ', () => {
  const { context } = loadGasScript();
  const currentText = [
    '# 旧タイトル',
    '',
    '- URL: https://old.example.com/a',
    '- 共有時のURL（短縮/リダイレクト元）: https://share.google/xyz',
    '- 保存日時: 2026-07-01 10:00',
    '- カテゴリ: PC系',
    '',
    '## 本文（自動抽出・参考）',
    '',
    'これは抽出された本文です。',
    ''
  ].join('\n');

  const rebuilt = context.rebuildDocContent_(
    currentText, '2026-07-01 10:00', 'PC系', '新タイトル', 'https://new.example.com/b', '新メモ'
  );

  assert.match(rebuilt, /^# 新タイトル$/m, 'タイトルが差し替わること');
  assert.match(rebuilt, /^- URL: https:\/\/new\.example\.com\/b$/m, 'URLが差し替わること');
  assert.match(rebuilt, /^- 共有時のURL（短縮\/リダイレクト元）: https:\/\/share\.google\/xyz$/m, '共有時URLを引き継ぐこと');
  assert.match(rebuilt, /^- 保存日時: 2026-07-01 10:00$/m, '保存日時は変えないこと');
  assert.match(rebuilt, /^- カテゴリ: PC系$/m, 'カテゴリは変えないこと');
  assert.match(rebuilt, /## メモ\n\n新メモ/, 'メモが差し替わること');
  assert.match(rebuilt, /## 本文（自動抽出・参考）\n\nこれは抽出された本文です。/, '自動抽出本文を引き継ぐこと');
  assert.doesNotMatch(rebuilt, /旧タイトル|旧メモ/, '旧の編集対象値が残らないこと');
});

test('編集API正常系: legacyな記事のfileIdで更新でき、Sheets行とDoc本文の両方に反映される', () => {
  const { context, sheet } = loadGasScript();
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'PC系', title: '元のタイトル', url: 'https://example.com/old', memo: '元メモ'
  });

  const listed = callDoGetList(context, 'PC系');
  assert.strictEqual(listed.items[0].fileId, fileId, '一覧APIが編集キーのfileIdを返すこと');

  const result = callDoPost(context, {
    action: 'update', fileId: fileId,
    title: '新タイトル', url: 'https://example.com/new', memo: '新メモ'
  });
  assert.strictEqual(result.ok, true);

  // 別レイヤー確認1: 一覧APIの返却値に反映されている
  const after = callDoGetList(context, 'PC系');
  assert.strictEqual(after.items[0].title, '新タイトル');
  assert.strictEqual(after.items[0].url, 'https://example.com/new');
  assert.strictEqual(after.items[0].memo, '新メモ');
  assert.strictEqual(after.items[0].fileId, fileId, 'fileIdは変わらないこと');

  // 別レイヤー確認2: Sheetsのタイトル列はHYPERLINK数式のまま新URLを指す
  assert.strictEqual(
    sheet._formulas[1][2],
    '=HYPERLINK("https://example.com/new","新タイトル")'
  );

  // 別レイヤー確認3: Googleドキュメント本文も更新される（保存日時・カテゴリは保持）
  const docText = context.DocumentApp.openById(fileId).getBody().getText();
  assert.match(docText, /^# 新タイトル$/m);
  assert.match(docText, /^- URL: https:\/\/example\.com\/new$/m);
  assert.match(docText, /^- カテゴリ: PC系$/m);
  assert.match(docText, /## メモ\n\n新メモ/);
});

test('編集API異常系: fileId無し・未知のfileId・タイトル空・URL形式不正はエラーになる', () => {
  const { context } = loadGasScript();
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'PC系', title: '元のタイトル', url: 'https://example.com/old'
  });

  const cases = [
    [{ action: 'update', title: 't', url: 'https://a.example/' }, '編集対象が指定されていません'],
    [{ action: 'update', fileId: 'doc-存在しない', title: 't', url: 'https://a.example/' }, '編集対象の記事が見つかりません'],
    [{ action: 'update', fileId: fileId, title: '', url: 'https://a.example/' }, 'タイトルを入力してください'],
    [{ action: 'update', fileId: fileId, title: 't', url: 'ftp://a.example/' }, 'URLの形式が不正です']
  ];
  for (const [body, message] of cases) {
    const result = callDoPost(context, body);
    assert.strictEqual(result.ok, false, JSON.stringify(body) + ' はエラーになること');
    assert.match(result.error, new RegExp(message));
  }
});

test('編集API認可: SHARED_TOKEN設定時、合言葉が違うupdateは拒否される', () => {
  const { context } = loadGasScript({
    scriptProperties: { SHARED_TOKEN: 'aikotoba' }
  });
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'PC系', title: '元のタイトル', url: 'https://example.com/old'
  });

  const denied = callDoPost(context, {
    action: 'update', fileId: fileId, title: '改ざん', url: 'https://evil.example/', token: 'ちがう'
  });
  assert.strictEqual(denied.ok, false);
  assert.match(denied.error, /合言葉が一致しません/);

  const allowed = callDoPost(context, {
    action: 'update', fileId: fileId, title: '正規の編集', url: 'https://example.com/new', token: 'aikotoba'
  });
  assert.strictEqual(allowed.ok, true);
});

test('編集API防御: メモの数式インジェクションはサニタイズされ、タイトルの引用符は数式内でエスケープされる', () => {
  const { context, sheet } = loadGasScript();
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'PC系', title: '元のタイトル', url: 'https://example.com/old'
  });

  const result = callDoPost(context, {
    action: 'update', fileId: fileId,
    title: '新"タイトル"', url: 'https://example.com/new', memo: '=IMPORTXML("https://evil.example/","//a")'
  });
  assert.strictEqual(result.ok, true);

  assert.strictEqual(
    sheet._rows[1][4], "'=IMPORTXML(\"https://evil.example/\",\"//a\")",
    'メモ先頭の = はアポストロフィで無害化されること'
  );
  assert.strictEqual(
    sheet._formulas[1][2],
    '=HYPERLINK("https://example.com/new","新""タイトル""")',
    'タイトル内の引用符は数式リテラル内でエスケープされること'
  );
});

// ---- 新規追加関数の重複定義チェック（教訓の再発防止パターン） ----------

test('Code.gs: 主要関数の定義がそれぞれちょうど1つである（重複定義の再発防止）', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.gs'), 'utf8');
  const names = [
    'handleReorderCategories_', 'getOrCreateRootFolder_', 'handleList_', 'getOrCreateIndexSheet_',
    'appendIndexRow_', 'doGet', 'doPost', 'handleUpdate_', 'extractFileIdFromFormula_',
    'rebuildDocContent_', 'extractDocOriginalUrl_', 'extractDocBodyText_',
    'getTags_', 'saveTags_', 'sanitizeTagName_', 'normalizeTagsInput_',
    'handleAddTag_', 'handleRemoveTag_', 'handleReorderTags_', 'fetchPageTitle_',
    'splitTagsText_', 'migrateRowIds_', 'findRowIndexById_', 'ensureIndexColumns_',
    'handleDelete_'
  ];
  for (const fn of names) {
    const definitions = source.match(new RegExp('function ' + fn + '\\(', 'g')) || [];
    assert.strictEqual(definitions.length, 1, fn + ' の定義がちょうど1つであること');
  }
});

// ============================================================
// ID列（行一意ID）のマイグレーションと、id優先の編集
// ============================================================

// ---- migrateRowIds_（冪等性・想定外状態） ----------------------

test('migrateRowIds_: ID列が空の既存行にUUIDが発行される', () => {
  const { context, sheet } = loadGasScript();
  // 移行前の7列（ID列なし）のシートを再現する
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '']);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '記事2', 'https://example.com/2', '', '', '']);

  context.getOrCreateIndexSheet_();

  assert.strictEqual(sheet._rows[0][7], 'ID', 'H1にID列のヘッダーが補完される');
  assert.ok(sheet._rows[1][7], '既存行1にIDが発行される');
  assert.ok(sheet._rows[2][7], '既存行2にIDが発行される');
  assert.notStrictEqual(sheet._rows[1][7], sheet._rows[2][7], '行ごとに異なるIDが振られる');
  assert.strictEqual(sheet._rows.length, 3, '行は増減しない');
});

test('migrateRowIds_: 何度実行してもID列の値が変わらない（冪等性）', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '']);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '記事2', 'https://example.com/2', '', '', '']);

  context.getOrCreateIndexSheet_();
  const firstPass = sheet._rows.map((r) => r[7]);

  // 実運用で一覧を複数回GETした状況に相当する（getOrCreateIndexSheet_ が毎回走る）
  context.getOrCreateIndexSheet_();
  context.getOrCreateIndexSheet_();

  assert.deepStrictEqual(sheet._rows.map((r) => r[7]), firstPass, '2回目以降の実行でIDが変化しない');
});

test('migrateRowIds_: 既存IDが数値型でも上書きせず温存する', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  // ID列がテキストではなく数値として保存されているシート（想定外状態）
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', 12345]);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '記事2', 'https://example.com/2', '', '', '', '']);

  context.getOrCreateIndexSheet_();

  assert.strictEqual(sheet._rows[1][7], 12345, '数値のIDは型ごと温存される');
  assert.ok(sheet._rows[2][7], '空の行にだけ新しいIDが発行される');
});

test('migrateRowIds_: 既存IDが重複していても修復しない（既存値優先）', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', 'dup-1']);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '記事2', 'https://example.com/2', '', '', '', 'dup-1']);

  context.getOrCreateIndexSheet_();

  assert.strictEqual(sheet._rows[1][7], 'dup-1', '重複IDは書き換えられない');
  assert.strictEqual(sheet._rows[2][7], 'dup-1', '重複IDは書き換えられない');
});

test('migrateRowIds_: データ行が無いシート（ヘッダーのみ）でも例外にならない', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);

  context.getOrCreateIndexSheet_(); // 例外が出ないこと自体が検証内容

  assert.strictEqual(sheet._rows.length, 1, 'ヘッダーのみのまま変化しない');
});

test('migrateRowIds_: 列を7列に切り詰めたシートでもID列を確保してから移行する', () => {
  // 余分な列を削除して運用しているシート（想定外状態）。
  // ID列（H列=8）が存在しないため、確保せずに書くと範囲外エラーになる
  const { context, sheet } = loadGasScript({ maxColumns: 7 });
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '']);

  context.getOrCreateIndexSheet_();

  assert.ok(sheet._columnCount >= 8, 'ID列ぶんの列が追加される');
  assert.strictEqual(sheet._rows[0][7], 'ID', 'ヘッダーが書き込める');
  assert.ok(sheet._rows[1][7], '既存行にIDが発行される');
});

test('migrateRowIds_: ヘッダ行が2行あるシートでも2行目はデータ行として扱われIDが振られる', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  // 誤って見出しが2行入っているシート（想定外状態）。行削除等の修復はしない方針
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', '']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', '']);

  context.getOrCreateIndexSheet_();

  assert.ok(sheet._rows[1][7], '2行目もデータ行としてIDが振られる（害はないため修復しない）');
  assert.ok(sheet._rows[2][7], '実データ行にもIDが振られる');
  assert.strictEqual(sheet._rows.length, 3, '行の削除・追加は行わない');
});

// ---- appendIndexRow_ / 一覧API ---------------------------------

test('save: 新規保存した行には必ずIDが付与され、一覧APIでも返る', () => {
  const { context, sheet } = loadGasScript({
    categories: ['カテゴリA'],
    fetchImpl: () => makeFetchResponse({ body: '<title>新規記事</title>' })
  });

  const result = callDoPost(context, { url: 'https://example.com/new', category: 'カテゴリA' });
  assert.strictEqual(result.ok, true);

  const id = sheet._rows[1][7];
  assert.ok(id, 'ID列に値が入る');

  const list = callDoGetList(context, 'カテゴリA');
  assert.strictEqual(list.ok, true);
  assert.strictEqual(list.items.length, 1);
  assert.strictEqual(list.items[0].id, id, '一覧APIが items[].id を返す');
});

test('save: 連続保存した行のIDは互いに異なる', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/1', category: 'カテゴリA' });
  callDoPost(context, { url: 'https://example.com/2', category: 'カテゴリA' });

  assert.ok(sheet._rows[1][7]);
  assert.ok(sheet._rows[2][7]);
  assert.notStrictEqual(sheet._rows[1][7], sheet._rows[2][7], '行ごとに一意なIDになる');
});

test('一覧API: 移行前の旧データ（ID列が空）でも、一覧取得時にIDが振られて返る', () => {
  const { context } = loadGasScript({ categories: ['カテゴリA'] });
  seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'カテゴリA',
    title: '旧記事', url: 'https://example.com/old', memo: ''
  });

  const list = callDoGetList(context, 'カテゴリA');

  assert.strictEqual(list.items.length, 1);
  assert.ok(list.items[0].id, 'マイグレーションで振られたIDが一覧に含まれる');
  assert.ok(list.items[0].fileId, '旧データは fileId も引き続き返る');
});

// ---- findRowIndexById_ -----------------------------------------

test('findRowIndexById_: 一致行の行番号を返し、未知のIDには -1 を返す', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', 'id-a']);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '記事2', 'https://example.com/2', '', '', '', 'id-b']);

  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), 'id-a'), 2);
  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), 'id-b'), 3);
  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), 'id-none'), -1);
});

test('findRowIndexById_: ID重複時は先頭（最古）の行を返す', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '古い方', 'https://example.com/1', '', '', '', 'dup']);
  sheet.appendRow(['2026-07-02 10:00', 'カテゴリA', '新しい方', 'https://example.com/2', '', '', '', 'dup']);

  assert.strictEqual(
    context.findRowIndexById_(sheet, sheet.getLastRow(), 'dup'), 2,
    '先頭の一致行を採用する'
  );
});

test('findRowIndexById_: 数値型で保存されたIDにも文字列比較で一致する', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', 12345]);

  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), '12345'), 2);
});

test('findRowIndexById_: 空ID・データ行なしでは -1 を返す（空セルへの誤一致防止）', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'カテゴリA', '記事1', 'https://example.com/1', '', '', '', '']);

  // ID列が空の行に対し空IDで検索しても一致してはいけない（全行が誤ヒットするのを防ぐ）
  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), ''), -1);
  assert.strictEqual(context.findRowIndexById_(sheet, sheet.getLastRow(), '   '), -1);
  assert.strictEqual(context.findRowIndexById_(sheet, 1, 'id-a'), -1, 'データ行が無ければ -1');
});

// ---- handleUpdate_（id優先・fileIdフォールバック） ---------------

test('update: id指定で該当行のSheetsが更新される', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id,
    title: '編集後タイトル', url: 'https://example.com/edited', memo: '編集後メモ'
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.id, id);
  assert.strictEqual(sheet._rows[1][2], '編集後タイトル', 'タイトル（HYPERLINK表示値）が更新される');
  assert.strictEqual(sheet._rows[1][3], 'https://example.com/edited');
  assert.strictEqual(sheet._rows[1][4], '編集後メモ');
  assert.strictEqual(sheet._rows[1][7], id, 'IDは編集で変化しない');
});

test('update: id指定は DocumentApp を一切呼ばない（新規記事はDocを持たないため）', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const id = sheet._rows[1][7];

  // DocumentApp.openById が呼ばれたら即座に失敗させる
  context.DocumentApp.openById = () => {
    throw new Error('id指定の編集で DocumentApp.openById が呼ばれてはいけない');
  };

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: ''
  });

  assert.strictEqual(result.ok, true, 'DocumentApp を呼ばずに完了すること');
});

test('update: fileIdのみ指定なら従来どおりGoogleドキュメント本文も更新される', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'カテゴリA',
    title: '旧記事', url: 'https://example.com/old', memo: '旧メモ'
  });

  const result = callDoPost(context, {
    action: 'update', fileId: fileId,
    title: '新タイトル', url: 'https://example.com/new', memo: '新メモ'
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][2], '新タイトル', 'Sheets行が更新される');

  const docText = context.DocumentApp.openById(fileId).getBody().getText();
  assert.match(docText, /# 新タイトル/, 'Doc本文のタイトルが更新される');
  assert.match(docText, /https:\/\/example\.com\/new/, 'Doc本文のURLが更新される');
  assert.match(docText, /新メモ/, 'Doc本文のメモが更新される');
});

test('update: id と fileId の両方が指定されたら id を優先しDocは触らない', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'カテゴリA',
    title: '旧記事', url: 'https://example.com/old', memo: '旧メモ'
  });
  // マイグレーションでこの旧行にもIDが振られる
  context.getOrCreateIndexSheet_();
  const id = sheet._rows[1][7];
  assert.ok(id, '前提: 旧行にIDが振られていること');
  const docBefore = context.DocumentApp.openById(fileId).getBody().getText();

  const result = callDoPost(context, {
    action: 'update', id: id, fileId: fileId,
    title: 'id優先で編集', url: 'https://example.com/byid', memo: ''
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][2], 'id優先で編集', 'Sheets行は更新される');
  assert.strictEqual(
    context.DocumentApp.openById(fileId).getBody().getText(), docBefore,
    'id優先のためDoc本文は変更されない'
  );
});

test('update: 未知のidはエラーになり、他の行を書き換えない', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const before = sheet._rows[1].slice();

  const result = callDoPost(context, {
    action: 'update', id: 'uuid-存在しない',
    title: '書き換え', url: 'https://evil.example/', memo: ''
  });

  assert.strictEqual(result.ok, false);
  assert.match(result.error, /編集対象の記事が見つかりません/);
  assert.deepStrictEqual(Array.from(sheet._rows[1]), Array.from(before), '既存行は変更されない');
});

test('update: id も fileId も無ければ「編集対象が指定されていません」', () => {
  const { context } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });

  const result = callDoPost(context, {
    action: 'update', title: 'タイトル', url: 'https://example.com/x', memo: ''
  });

  assert.strictEqual(result.ok, false);
  assert.match(result.error, /編集対象が指定されていません/);
});

test('update: id・fileId が空文字だけでも「編集対象が指定されていません」', () => {
  const { context } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });

  // PWA側は id/fileId を常に送るため、両方空文字のケースを明示的に守る
  const result = callDoPost(context, {
    action: 'update', id: '', fileId: '', title: 'タイトル', url: 'https://example.com/x', memo: ''
  });

  assert.strictEqual(result.ok, false);
  assert.match(result.error, /編集対象が指定されていません/);
});

test('update: id が空文字で fileId のみ有効なら従来のDoc更新経路になる', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'カテゴリA',
    title: '旧記事', url: 'https://example.com/old', memo: '旧メモ'
  });

  // PWA側が id: '' を常に送る実装のため、空文字が優先キーとして誤採用されないことを守る
  const result = callDoPost(context, {
    action: 'update', id: '', fileId: fileId,
    title: '新タイトル', url: 'https://example.com/new', memo: ''
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][2], '新タイトル');
  assert.match(
    context.DocumentApp.openById(fileId).getBody().getText(), /# 新タイトル/,
    'fileIdフォールバックとしてDoc本文も更新される'
  );
});

test('update: id指定でも入力検証（タイトル必須・URL形式）は従来どおり働く', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const id = sheet._rows[1][7];

  const noTitle = callDoPost(context, {
    action: 'update', id: id, title: '', url: 'https://example.com/x', memo: ''
  });
  assert.strictEqual(noTitle.ok, false);
  assert.match(noTitle.error, /タイトルを入力してください/);

  const badUrl = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'ftp://example.com/x', memo: ''
  });
  assert.strictEqual(badUrl.ok, false);
  assert.match(badUrl.error, /URLの形式が不正です/);
});

test('update: id指定でもメモの数式インジェクションは無害化される', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル',
    url: 'https://example.com/x', memo: '=IMPORTXML("https://evil.example/","//a")'
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(
    sheet._rows[1][4], "'=IMPORTXML(\"https://evil.example/\",\"//a\")",
    'メモ先頭の = はアポストロフィで無害化されること'
  );
});

test('update: id指定でも SHARED_TOKEN 設定時は合言葉が照合される', () => {
  const { context, sheet } = loadGasScript({
    categories: ['カテゴリA'], scriptProperties: { SHARED_TOKEN: 'aikotoba' }
  });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA', token: 'aikotoba' });
  const id = sheet._rows[1][7];

  const denied = callDoPost(context, {
    action: 'update', id: id, title: '改ざん', url: 'https://evil.example/', token: 'ちがう'
  });
  assert.strictEqual(denied.ok, false);
  assert.match(denied.error, /合言葉が一致しません/);
  assert.notStrictEqual(sheet._rows[1][2], '改ざん', '拒否時は行が書き換わらない');
});

// ---- handleUpdate_ の tags 対応（Phase4） -------------------------

test('update: tagsを配列で指定するとタグ列が置き換わる', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  callDoPost(context, { action: 'addTag', name: 'stock2' });
  callDoPost(context, {
    url: 'https://example.com/orig', category: 'カテゴリA', tags: ['stock1']
  });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: '',
    tags: ['stock2']
  });

  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.tags), ['stock2']);
  assert.strictEqual(sheet._rows[1][6], 'stock2', 'タグ列がstock2のみに置き換わる');
});

test('update: tags未指定（配列でない）ならタグ列は現状を保持する', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  callDoPost(context, {
    url: 'https://example.com/orig', category: 'カテゴリA', tags: ['stock1']
  });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: ''
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.tags, null, '未指定時はtagsをnullで返す');
  assert.strictEqual(sheet._rows[1][6], 'stock1', 'タグ列は変更されず保持される');
});

test('update: 空配列を指定するとタグ列が全クリアされる', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  callDoPost(context, {
    url: 'https://example.com/orig', category: 'カテゴリA', tags: ['stock1']
  });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: '',
    tags: []
  });

  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.tags), []);
  assert.strictEqual(sheet._rows[1][6], '', 'タグ列が空になる');
});

test('update: 未登録タグを指定するとエラーになり、行が書き換わらない', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  callDoPost(context, {
    url: 'https://example.com/orig', category: 'カテゴリA', tags: ['stock1']
  });
  const id = sheet._rows[1][7];
  const before = sheet._rows[1].slice();

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: '',
    tags: ['stock1', '未登録タグ']
  });

  assert.strictEqual(result.ok, false);
  assert.match(result.error, /不明なタグです/);
  assert.deepStrictEqual(Array.from(sheet._rows[1]), Array.from(before), '既存行は変更されない');
});

test('update: tagsの前後空白は正規化されてから照合・保存される', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  callDoPost(context, { url: 'https://example.com/orig', category: 'カテゴリA' });
  const id = sheet._rows[1][7];

  const result = callDoPost(context, {
    action: 'update', id: id, title: 'タイトル', url: 'https://example.com/edited', memo: '',
    tags: ['', '  stock1  ']
  });

  assert.strictEqual(result.ok, true);
  assert.deepStrictEqual(Array.from(result.tags), ['stock1'], '空文字は除外・前後空白はtrimされること');
  assert.strictEqual(sheet._rows[1][6], 'stock1');
});

test('update: fileIdフォールバック経路でもtagsが指定されれば反映される', () => {
  const { context, sheet } = loadGasScript({ categories: ['カテゴリA'] });
  callDoPost(context, { action: 'addTag', name: 'stock1' });
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'カテゴリA',
    title: '旧記事', url: 'https://example.com/old', memo: '旧メモ'
  });

  const result = callDoPost(context, {
    action: 'update', fileId: fileId,
    title: '新タイトル', url: 'https://example.com/new', memo: '', tags: ['stock1']
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows[1][6], 'stock1');
});

// ============================================================
// 記事の削除（action=delete）
// ============================================================

test('削除API: 存在するidを削除すると行数が1減り、他の行は変わらない', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: (url) => makeFetchResponse({ body: '<title>記事' + url.split('/').pop() + '</title>' })
  });
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系', memo: 'メモA' });
  callDoPost(context, { url: 'https://example.com/b', category: 'PC系', memo: 'メモB' });

  const listed = callDoGetList(context, 'PC系');
  const target = listed.items[1]; // 古い方（記事a）を削除する
  assert.strictEqual(target.url, 'https://example.com/a');
  const rowCountBefore = sheet._rows.length;

  const result = callDoPost(context, { action: 'delete', id: target.id });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.id, target.id);

  assert.strictEqual(sheet._rows.length, rowCountBefore - 1, '行数が1減る');
  const after = callDoGetList(context, 'PC系');
  assert.strictEqual(after.items.length, 1, '一覧からも消える');
  assert.strictEqual(after.items[0].title, '記事b', '残った行は変更されない');
  assert.strictEqual(after.items[0].memo, 'メモB');
});

test('削除API: 存在しないidは「削除対象の記事が見つかりません」エラーになり、行は変更されない', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系' });
  const before = sheet._rows.map((r) => r.slice());

  const result = callDoPost(context, { action: 'delete', id: 'uuid-存在しない' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /削除対象の記事が見つかりません/);
  assert.deepStrictEqual(sheet._rows, before, 'どの行も変更されない');
});

test('削除API: idが指定されていない場合は「削除対象が指定されていません」エラー', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'delete' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /削除対象が指定されていません/);
});

test('削除API: idが空白のみの場合も「削除対象が指定されていません」エラー（trim後に空判定）', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'delete', id: '   ' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /削除対象が指定されていません/);
});

test('削除API: idの前後に空白があっても正しく一致してtrim後の値で削除される', () => {
  const { context, sheet } = loadGasScript();
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系' });
  const listed = callDoGetList(context, 'PC系');
  const id = listed.items[0].id;

  const result = callDoPost(context, { action: 'delete', id: '  ' + id + '  ' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows.length, 1, 'ヘッダのみになる');
});

test('削除API: データ行が1件も無い状態での削除は「削除対象の記事が見つかりません」エラー', () => {
  const { context } = loadGasScript();
  const result = callDoPost(context, { action: 'delete', id: 'uuid-1' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /削除対象の記事が見つかりません/);
});

test('削除API: token不一致は「合言葉が一致しません」エラーになり、行は削除されない', () => {
  const { context, sheet } = loadGasScript({ scriptProperties: { SHARED_TOKEN: 'aikotoba' } });
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系', token: 'aikotoba' });
  const listed = callDoGetList(context, 'PC系', 'aikotoba');
  const id = listed.items[0].id;
  const before = sheet._rows.map((r) => r.slice());

  const result = callDoPost(context, { action: 'delete', id: id, token: 'ちがう' });
  assert.strictEqual(result.ok, false);
  assert.match(result.error, /合言葉が一致しません/);
  assert.deepStrictEqual(sheet._rows, before, '行は削除されない');
});

test('削除API: Drive上のGoogleドキュメント（fileId列）は削除されない（非破壊）', () => {
  const { context, sheet } = loadGasScript();
  const { fileId } = seedLegacyArticle(context, {
    savedAt: '2026-07-01 10:00', category: 'PC系', title: '旧記事', url: 'https://example.com/old', memo: ''
  });

  // 一覧取得でマイグレーションが走り、legacy行にもidが付く
  const listed = callDoGetList(context, 'PC系');
  const id = listed.items[0].id;

  const result = callDoPost(context, { action: 'delete', id: id });
  assert.strictEqual(result.ok, true);

  const after = callDoGetList(context, 'PC系');
  assert.strictEqual(after.items.length, 0, 'Sheets行は削除される');
  // DocumentApp.openByIdが例外を投げなければDocは残っている（スタブは削除時にdocsByIdから除去しない実装のため）
  assert.doesNotThrow(() => context.DocumentApp.openById(fileId), 'Drive上のDocsは削除されず開けたままである');
});

test('削除API: ID重複行がある場合は先頭（最古）の一致行だけが削除される', () => {
  const { context, sheet } = loadGasScript();
  sheet.appendRow(['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID']);
  sheet.appendRow(['2026-07-01 10:00', 'PC系', '重複1', 'https://example.com/1', '', '', '', 'dup-id']);
  sheet.appendRow(['2026-07-02 10:00', 'PC系', '重複2', 'https://example.com/2', '', '', '', 'dup-id']);

  const result = callDoPost(context, { action: 'delete', id: 'dup-id' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet._rows.length, 2, 'ヘッダ+重複2のみ残る（2行）');
  assert.strictEqual(sheet._rows[1][2], '重複2', '後続の重複行は残る');
});

test('削除API: 並行編集で対象より前の行が別途削除され行番号がずれても、idで正しい行を特定して削除する', () => {
  const { context, sheet } = loadGasScript({
    fetchImpl: (url) => makeFetchResponse({ body: '<title>記事' + url.split('/').pop() + '</title>' })
  });
  callDoPost(context, { url: 'https://example.com/a', category: 'PC系' });
  callDoPost(context, { url: 'https://example.com/b', category: 'PC系' });
  callDoPost(context, { url: 'https://example.com/c', category: 'PC系' });

  const listed = callDoGetList(context, 'PC系');
  const idA = listed.items[2].id; // 記事a（最古・2行目）
  const idC = listed.items[0].id; // 記事c（最新・4行目）

  // 別クライアントが先に記事aを削除する（=クライアントBが記事cのidを取得した時点の
  // 行番号想定と、実際の削除実行時点でずれが生じる状況を再現）。
  const first = callDoPost(context, { action: 'delete', id: idA });
  assert.strictEqual(first.ok, true, '記事aの削除が先に成功する');
  assert.strictEqual(sheet._rows.length, 3, 'ヘッダ+2行（b, c）に減る');

  // クライアントBは古い行番号を覚えず、idCで再度特定して削除する（実装が
  // handleUpdate_と同じ「毎回idで走査し直す」方式であることの検証）。
  const second = callDoPost(context, { action: 'delete', id: idC });
  assert.strictEqual(second.ok, true, '行番号がずれてもidで記事cを正しく削除できる');

  const after = callDoGetList(context, 'PC系');
  assert.strictEqual(after.items.length, 1, '記事bのみ残る');
  assert.strictEqual(after.items[0].title, '記事b');
});

