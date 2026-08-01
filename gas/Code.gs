// ============================================================
// ナレッジ保存 バックエンド (Google Apps Script Webアプリ)
//
//   PWA(共有シート)から URL とカテゴリ名を POST で受け取り、
//   ページタイトルを自動取得して検索用インデックス（Sheets）に記録する。
//
//   このスクリプトは検索用インデックス・スプレッドシート「ナレッジ一覧」に
//   コンテナバインドされている前提（SpreadsheetApp.getActiveSpreadsheet()で
//   そのスプレッドシート自身を参照する）。
//
//   デプロイ方法: バインド先スプレッドシートの「拡張機能」→「Apps Script」→
//     「デプロイ」→「ウェブアプリ」
//     - 次のユーザーとして実行: 自分
//     - アクセスできるユーザー: 全員
//   ※「全員」で公開するため、スクリプトプロパティ SHARED_TOKEN に
//     合言葉を設定し、PWA側の設定画面にも同じ値を入れること。
// ============================================================

// ---- 定数 --------------------------------------------------

// Drive 上のルートフォルダ名（過去に保存したGoogleドキュメントの置き場。
// 新規記事はここには保存しないが、Driveショートカットボタンの参照先として使う）
var ROOT_FOLDER_NAME = 'ナレッジ';

// カテゴリ一覧を保存するスクリプトプロパティのキー
// （PWAから追加・削除できるようにするため、固定配列ではなくプロパティで管理する）
var CATEGORIES_PROPERTY = 'CATEGORIES_JSON';

// タグ一覧を保存するスクリプトプロパティのキー（カテゴリとは別軸の項目）
// 初期値は空配列。設定画面から追加していく運用とする
var TAGS_PROPERTY = 'TAGS_JSON';

// ルートフォルダ「ナレッジ」のIDを保存するスクリプトプロパティのキー
// （毎回マイドライブ全体から名前検索すると数百ms〜数秒かかるため、
//   一度見つけたフォルダのIDを覚えておき、以後はIDで直接開いて高速化する）
var ROOT_FOLDER_ID_PROPERTY = 'ROOT_FOLDER_ID';

// 初回アクセス時の初期カテゴリ一覧（Driveフォルダ名もカテゴリ名と同一にする）。
// 空配列が正常な初期状態で、設定画面（⚙）からユーザー自身が追加していく運用とする。
var DEFAULT_CATEGORIES = [];

// タイムゾーン（保存日時に使用）
var TIME_ZONE = 'Asia/Tokyo';

// 検索用インデックス（Googleスプレッドシート）関連の定数
// このスクリプトは「ナレッジ一覧」スプレッドシートにコンテナバインドされている前提のため、
// スプレッドシートIDをスクリプトプロパティに保持する必要はなく、
// SpreadsheetApp.getActiveSpreadsheet() で自分自身（バインド先）を直接参照する。
var INDEX_SHEET_NAME = 'ナレッジ一覧';
// ID列は行を一意に識別するUUID（migrateRowIds_ が既存行へ後から付与する）。
// 列は必ず末尾に足すこと。途中に挿入すると既存シートの全データを物理的にずらす必要があり、
// 破壊的な移行になるため（末尾追加なら既存列の位置は変わらない）。
var INDEX_HEADER = ['日時', 'カテゴリ', 'タイトル', 'URL', 'メモ', 'Driveファイル', 'タグ', 'ID'];
var INDEX_COL = { SAVED_AT: 1, CATEGORY: 2, TITLE: 3, URL: 4, MEMO: 5, FILE: 6, TAGS: 7, ID: 8 };

// 一覧APIで一度に返す件数（1ページ分）。
// 件数が増えてもレスポンスが重くならないよう50件で区切り、
// 続きは offset パラメータで取得する（PWA側の「さらに読み込む」ボタン用）
var LIST_PAGE_SIZE = 50;

// ---- エントリポイント ---------------------------------------

/**
 * GET エンドポイント。
 * - パラメータ無し: 動作確認用メッセージを返す（合言葉不要の稼働確認）
 * - ?action=list&category=xxx&token=xxx : そのカテゴリの保存済み記事一覧を返す
 * - ?action=categories&token=xxx : カテゴリ一覧とDriveの「ナレッジ」フォルダURLを返す
 * データを返す action は、POSTと同様に SHARED_TOKEN（設定時）の照合を必須とする。
 * 記事一覧はタイトル・URL・メモといった個人の閲覧記録に近い情報を含むため、
 * WebアプリURLを知られただけでは読めないようにしておく。
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var token = String((e && e.parameter && e.parameter.token) || '');
  if (action === 'list') {
    return handleList_(e.parameter.category, token, e.parameter.offset);
  }
  if (action === 'categories') {
    return handleCategories_(token);
  }
  return jsonResponse_({
    ok: true,
    message: 'ナレッジ保存APIは稼働中です。保存は POST で行います。'
  });
}

/**
 * カテゴリ一覧・タグ一覧と、Driveの「ナレッジ」フォルダを直接開くためのURLを返す。
 * PWAの起動時・設定画面・一覧画面でカテゴリ/タグボタンを動的に描画するために使う。
 */
function handleCategories_(token) {
  try {
    checkToken_(token);
    return jsonResponse_({
      ok: true,
      categories: getCategories_(),
      tags: getTags_(),
      rootFolderUrl: getOrCreateRootFolder_().getUrl()
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String((err && err.message) || err) });
  }
}

/**
 * 指定カテゴリの保存済み記事一覧を、検索用インデックス（Sheets）から新しい順に返す。
 * 一度に返すのは LIST_PAGE_SIZE 件まで。offset（新しい順で何件目から）を指定すると
 * 続きのページを返し、まだ続きがある場合はレスポンスに hasMore: true を含める。
 */
function handleList_(category, token, offsetParam) {
  try {
    checkToken_(token);
    if (!category || getCategories_().indexOf(category) === -1) {
      throw new Error('不明なカテゴリです: ' + category);
    }

    // offset の検証（未指定・不正値は 0 = 先頭ページとして扱う）
    var offset = parseInt(offsetParam, 10);
    if (isNaN(offset) || offset < 0) {
      offset = 0;
    }

    var sheet = getOrCreateIndexSheet_();
    var values = sheet.getDataRange().getValues(); // values[0] はヘッダ行

    // Driveファイル列の数式（=HYPERLINK("…/document/d/{fileId}/edit","開く")）を取得し、
    // 各記事の編集キーとなる fileId を抽出できるようにする。
    // getValues では数式セルは表示値（"開く"）になり fileId が取れないため、別途 getFormulas で取る。
    // fileFormulas[k] がシートの (k+2) 行目（＝ values[k+1]）に対応する。
    var fileFormulas = values.length >= 2
      ? sheet.getRange(2, INDEX_COL.FILE, values.length - 1, 1).getFormulas()
      : [];

    var items = [];
    var matched = 0;   // このカテゴリで何件目まで見たか（offsetの読み飛ばし用）
    var hasMore = false;
    // 新しい順（末尾の行から）に走査し、offset分を読み飛ばして1ページ分集める
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (row[INDEX_COL.CATEGORY - 1] !== category) {
        continue;
      }
      matched++;
      if (matched <= offset) {
        continue; // 前のページで返却済み
      }
      if (items.length >= LIST_PAGE_SIZE) {
        hasMore = true; // 1ページ分を超える該当行がまだある
        break;
      }
      var fileFormula = fileFormulas[i - 1] ? fileFormulas[i - 1][0] : '';
      items.push({
        // id は編集・削除時の行特定キー。fileId は旧データ向けのフォールバック
        id: String(row[INDEX_COL.ID - 1] == null ? '' : row[INDEX_COL.ID - 1]).trim(),
        savedAt: row[INDEX_COL.SAVED_AT - 1],
        title: row[INDEX_COL.TITLE - 1],
        url: row[INDEX_COL.URL - 1],
        memo: row[INDEX_COL.MEMO - 1] || '',
        tags: splitTagsText_(row[INDEX_COL.TAGS - 1]),
        fileId: extractFileIdFromFormula_(fileFormula)
      });
    }

    return jsonResponse_({
      ok: true, category: category, items: items, offset: offset, hasMore: hasMore
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String((err && err.message) || err) });
  }
}

/** インデックスシートのタグ列（", " 区切り文字列）を、前後空白除去済みの配列に戻す */
function splitTagsText_(tagsText) {
  if (!tagsText) { return []; }
  return String(tagsText).split(',').map(function (t) { return t.trim(); }).filter(function (t) { return t; });
}

/**
 * PWA からのPOSTリクエストを受け取るエントリポイント。
 * body.action で処理を振り分ける（省略時は 'save' = 記事の保存）。
 *   - save             : { url, category, memo? } を保存
 *   - addCategory      : { name } をカテゴリ一覧に追加
 *   - removeCategory   : { name } をカテゴリ一覧から削除（保存済みデータは残す）
 *   - reorderCategories: { categories } の順にカテゴリの並び順を変更
 *   - update           : { id?, fileId?, title, url, memo?, tags? } 保存済み記事の内容を編集
 *   - addTag           : { name } をタグ一覧に追加
 *   - removeTag        : { name } をタグ一覧から削除（保存済みデータは残す）
 *   - reorderTags      : { tags } の順にタグの並び順を変更
 * すべての action 共通で token（合言葉）を検証する。
 */
function doPost(e) {
  try {
    var body = parseJsonBody_(e);
    checkToken_(String(body.token || ''));

    var action = String(body.action || 'save');
    switch (action) {
      case 'save':
        return jsonResponse_(handleSave_(body));
      case 'addCategory':
        return jsonResponse_(handleAddCategory_(body));
      case 'removeCategory':
        return jsonResponse_(handleRemoveCategory_(body));
      case 'reorderCategories':
        return jsonResponse_(handleReorderCategories_(body));
      case 'update':
        return jsonResponse_(handleUpdate_(body));
      case 'addTag':
        return jsonResponse_(handleAddTag_(body));
      case 'removeTag':
        return jsonResponse_(handleRemoveTag_(body));
      case 'reorderTags':
        return jsonResponse_(handleReorderTags_(body));
      default:
        throw new Error('不明なactionです: ' + action);
    }
  } catch (err) {
    // エラー内容を日本語メッセージで返す（PWA側で表示する）
    return jsonResponse_({
      ok: false,
      error: String((err && err.message) || err)
    });
  }
}

// ---- リクエスト処理 -----------------------------------------

/**
 * POST ボディのJSONを解析する。不正な場合は日本語メッセージ付きの Error を投げる。
 */
function parseJsonBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('リクエストボディがありません');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (parseErr) {
    throw new Error('リクエストボディのJSONが不正です');
  }
}

/**
 * 記事保存(action=save)のリクエストを検証し、保存処理を実行する。
 * 保存先は検索用インデックス（Sheets）のみ（Googleドキュメントへの転記は行わない）。
 */
function handleSave_(body) {
  var params = validateSaveParams_(body);

  // ステップ1: 共有された URL を実URLへ解決する
  // （Androidの共有機能が生成する share.google 等の短縮/リダイレクトリンクのままだと、
  //   後で人間やAIが開く際に不便な上、短縮リンクは将来失効するリスクもあるため）
  var resolvedUrl = resolveFinalUrl_(params.url);

  // ステップ2: ページのタイトルを取得（失敗時はタイトル=URL）
  var title = fetchPageTitle_(resolvedUrl);

  // ステップ3: メモがタイトルと重複している場合は捨てる
  // （Android共有時に「タイトル文字列」がそのままメモ扱いで送られてくることが多く、
  //   タイトルと同じ内容が二重表示されるのを防ぐ）
  var memo = isDuplicateMemo_(params.memo, title) ? '' : params.memo;

  // ステップ4: 検索用インデックス（Sheets）に1行追記する
  var savedAt = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm');
  appendIndexRow_(savedAt, params.category, title, resolvedUrl, memo, params.tags);

  return { ok: true, title: title, category: params.category, tags: params.tags };
}

/**
 * save リクエストの必須項目を検証して返す。不正な場合は日本語メッセージ付きの Error を投げる。
 */
function validateSaveParams_(body) {
  var url = String(body.url || '').trim();
  var category = String(body.category || '').trim();
  var memo = String(body.memo || '').trim();
  var tags = normalizeTagsInput_(body.tags);

  // URL の検証: http/https のみ許可（javascript: 等の混入を防ぐ）
  if (!url) {
    throw new Error('URLが指定されていません');
  }
  if (!/^https?:\/\/\S+$/i.test(url)) {
    throw new Error('URLの形式が不正です: ' + url);
  }

  // カテゴリの検証: 現在のカテゴリ一覧に完全一致するもののみ受け付ける
  if (!category) {
    throw new Error('カテゴリが指定されていません');
  }
  if (getCategories_().indexOf(category) === -1) {
    throw new Error('不明なカテゴリです: ' + category);
  }

  // タグの検証: 現在のタグ一覧に完全一致するもののみ受け付ける（任意項目のため0件でも可）
  var knownTags = getTags_();
  tags.forEach(function (tag) {
    if (knownTags.indexOf(tag) === -1) {
      throw new Error('不明なタグです: ' + tag);
    }
  });

  return { url: url, category: category, memo: memo, tags: tags };
}

/**
 * リクエストの tags（配列であるべき）を、前後空白除去済み・空文字除外済みの
 * 文字列配列に正規化する純粋関数。配列でない場合は空配列を返す。
 */
function normalizeTagsInput_(raw) {
  if (Object.prototype.toString.call(raw) !== '[object Array]') {
    return [];
  }
  return raw
    .map(function (tag) { return String(tag || '').trim(); })
    .filter(function (tag) { return !!tag; });
}

/**
 * カテゴリを1件追加する(action=addCategory)。同名が既にあればエラー。
 */
function handleAddCategory_(body) {
  var name = sanitizeCategoryName_(body.name);
  if (!name) {
    throw new Error('カテゴリ名が指定されていません');
  }

  var categories = getCategories_();
  if (categories.indexOf(name) !== -1) {
    throw new Error('同名のカテゴリが既にあります: ' + name);
  }

  categories.push(name);
  saveCategories_(categories);
  return { ok: true, categories: categories };
}

/**
 * カテゴリを1件削除する(action=removeCategory)。
 * カテゴリ一覧（選択肢）から外すだけで、Drive上の保存済みファイルや
 * Sheetsの過去の行は削除しない（非破壊。改名は別カテゴリの削除+追加で代用する）。
 */
function handleRemoveCategory_(body) {
  var name = String(body.name || '').trim();
  var categories = getCategories_();
  var index = categories.indexOf(name);
  if (index === -1) {
    throw new Error('存在しないカテゴリです: ' + name);
  }

  categories.splice(index, 1);
  saveCategories_(categories);
  return { ok: true, categories: categories };
}

/**
 * カテゴリの並び順を変更する(action=reorderCategories)。
 * body.categories（新しい並び順の配列）が「現在のカテゴリ一覧の並び替え」に
 * なっていることを検証してから保存する（追加・削除・改名の混入を防ぐ）。
 * 並び順はカテゴリ一覧そのもの（配列の順序）として永続化されるため、
 * 保存画面・一覧画面・設定画面のすべてに同じ順序が反映される。
 */
function handleReorderCategories_(body) {
  var requested = body.categories;
  if (Object.prototype.toString.call(requested) !== '[object Array]') {
    throw new Error('並び替え後のカテゴリ一覧が指定されていません');
  }

  var normalized = requested.map(function (name) { return String(name); });
  var current = getCategories_();

  // 要素の集合が完全一致するか（順序だけの違いか）をソート済みJSONで比較する
  var sortedRequested = JSON.stringify(normalized.slice().sort());
  var sortedCurrent = JSON.stringify(current.slice().sort());
  if (normalized.length !== current.length || sortedRequested !== sortedCurrent) {
    throw new Error('並び替えの内容が現在のカテゴリ一覧と一致しません。画面を開き直してからやり直してください');
  }

  saveCategories_(normalized);
  return { ok: true, categories: normalized };
}

/**
 * カテゴリ名として使える形に整える純粋関数。
 * Driveのフォルダ名としても使うため、区切り文字と誤認されうる / \ を置換する。
 */
function sanitizeCategoryName_(name) {
  return String(name || '').trim().replace(/[\\/]/g, '・');
}

/**
 * タグを1件追加する(action=addTag)。同名が既にあればエラー。
 */
function handleAddTag_(body) {
  var name = sanitizeTagName_(body.name);
  if (!name) {
    throw new Error('タグ名が指定されていません');
  }

  var tags = getTags_();
  if (tags.indexOf(name) !== -1) {
    throw new Error('同名のタグが既にあります: ' + name);
  }

  tags.push(name);
  saveTags_(tags);
  return { ok: true, tags: tags };
}

/**
 * タグを1件削除する(action=removeTag)。
 * タグ一覧（選択肢）から外すだけで、Sheetsの過去の行に書き込まれたタグ文字列は変更しない。
 */
function handleRemoveTag_(body) {
  var name = String(body.name || '').trim();
  var tags = getTags_();
  var index = tags.indexOf(name);
  if (index === -1) {
    throw new Error('存在しないタグです: ' + name);
  }

  tags.splice(index, 1);
  saveTags_(tags);
  return { ok: true, tags: tags };
}

/**
 * タグの並び順を変更する(action=reorderTags)。
 * body.tags（新しい並び順の配列）が「現在のタグ一覧の並び替え」になっていることを
 * 検証してから保存する（カテゴリのreorderCategoriesと同じ方針）。
 */
function handleReorderTags_(body) {
  var requested = body.tags;
  if (Object.prototype.toString.call(requested) !== '[object Array]') {
    throw new Error('並び替え後のタグ一覧が指定されていません');
  }

  var normalized = requested.map(function (name) { return String(name); });
  var current = getTags_();

  var sortedRequested = JSON.stringify(normalized.slice().sort());
  var sortedCurrent = JSON.stringify(current.slice().sort());
  if (normalized.length !== current.length || sortedRequested !== sortedCurrent) {
    throw new Error('並び替えの内容が現在のタグ一覧と一致しません。画面を開き直してからやり直してください');
  }

  saveTags_(normalized);
  return { ok: true, tags: normalized };
}

/**
 * タグ名として使える形に整える純粋関数。
 * Sheetsのタグ列はカンマ区切りで1セルにまとめるため、区切り文字と誤認されうる , を置換する。
 */
function sanitizeTagName_(name) {
  return String(name || '').trim().replace(/,/g, '、');
}

/**
 * スクリプトプロパティ SHARED_TOKEN が設定されている場合、
 * リクエストの token と一致するか確認する。未設定なら素通し。
 */
function checkToken_(token) {
  var expected = '';
  try {
    expected = PropertiesService.getScriptProperties().getProperty('SHARED_TOKEN') || '';
  } catch (e) {
    expected = '';
  }
  if (expected && token !== expected) {
    throw new Error('合言葉が一致しません。PWAの設定画面を確認してください');
  }
}

// ---- カテゴリ一覧の永続化（スクリプトプロパティ） --------------

/**
 * 現在のカテゴリ一覧を返す。未初期化・壊れている場合は初期カテゴリ（空配列）で初期化してから返す。
 * タグと同様、初期値は空配列（0件）自体が正常な状態のため、
 * 「保存されていた配列が空かどうか」では再初期化を判断しない。
 */
function getCategories_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(CATEGORIES_PROPERTY);

  if (!raw) {
    saveCategories_(DEFAULT_CATEGORIES.slice());
    return DEFAULT_CATEGORIES.slice();
  }

  try {
    var list = JSON.parse(raw);
    if (Object.prototype.toString.call(list) === '[object Array]') {
      return list;
    }
  } catch (parseErr) {
    // 壊れていた場合は初期カテゴリへフォールバック（下で再初期化する）
  }

  saveCategories_(DEFAULT_CATEGORIES.slice());
  return DEFAULT_CATEGORIES.slice();
}

/** カテゴリ一覧をスクリプトプロパティへ保存する。 */
function saveCategories_(categories) {
  PropertiesService.getScriptProperties().setProperty(CATEGORIES_PROPERTY, JSON.stringify(categories));
}

// ---- タグ一覧の永続化（スクリプトプロパティ） ------------------

/**
 * 現在のタグ一覧を返す。未初期化・壊れている場合は空配列で初期化してから返す。
 * カテゴリと異なり初期値は空配列（0件）自体が正常な状態のため、
 * 「保存されていた配列が空かどうか」では再初期化を判断しない。
 */
function getTags_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(TAGS_PROPERTY);

  if (!raw) {
    saveTags_([]);
    return [];
  }

  try {
    var list = JSON.parse(raw);
    if (Object.prototype.toString.call(list) === '[object Array]') {
      return list;
    }
  } catch (parseErr) {
    // 壊れていた場合は空配列へフォールバック（下で再初期化する）
  }

  saveTags_([]);
  return [];
}

/** タグ一覧をスクリプトプロパティへ保存する。 */
function saveTags_(tags) {
  PropertiesService.getScriptProperties().setProperty(TAGS_PROPERTY, JSON.stringify(tags));
}

// ---- URL解決 --------------------------------------------------

// 一般的なブラウザに近い UA を名乗る（ボット扱いで拒否されるサイト対策）
var DEFAULT_USER_AGENT = 'Mozilla/5.0 (Linux; Android 14) KnowledgeShareBot/1.0';

// リダイレクト解決の最大ホップ数（無限リダイレクト対策）
var MAX_REDIRECT_HOPS = 5;

/**
 * 短縮/リダイレクトリンク（share.google 等）を実際の記事URLへ解決する。
 * 3xxレスポンスの Location ヘッダを最大 MAX_REDIRECT_HOPS 回まで辿る。
 * 解決できなければ（通信エラー・Locationヘッダ無し等）、その時点のURLを返す。
 */
function resolveFinalUrl_(url) {
  var currentUrl = url;
  for (var i = 0; i < MAX_REDIRECT_HOPS; i++) {
    var response;
    try {
      response = UrlFetchApp.fetch(currentUrl, {
        muteHttpExceptions: true,
        followRedirects: false,
        headers: { 'User-Agent': DEFAULT_USER_AGENT }
      });
    } catch (fetchErr) {
      return currentUrl;
    }

    var code = response.getResponseCode();
    if (code < 300 || code >= 400) {
      return currentUrl;
    }

    var headers = response.getHeaders();
    var location = headers['Location'] || headers['location'];
    if (!location) {
      return currentUrl;
    }
    currentUrl = resolveRelativeUrl_(currentUrl, location);
  }
  return currentUrl;
}

/**
 * Location ヘッダの値（絶対URLとは限らない）を、遷移元URLを基準に絶対URLへ直す純粋関数。
 */
function resolveRelativeUrl_(baseUrl, location) {
  if (/^https?:\/\//i.test(location)) {
    return location;
  }
  var m = baseUrl.match(/^(https?:\/\/[^/]+)/i);
  var origin = m ? m[1] : '';
  if (location.charAt(0) === '/') {
    return origin + location;
  }
  return origin + '/' + location;
}

// ---- タイトル取得 -----------------------------------------------

/**
 * URL 先のページからタイトルを取得する。
 * <title> → og:title の順、どちらも取れなければ URL をタイトル代わりに使う。
 * 通信エラー・タイムアウト等はすべて握りつぶし、URL をタイトル代わりに返す。
 */
function fetchPageTitle_(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': DEFAULT_USER_AGENT }
    });

    if (response.getResponseCode() >= 400) {
      return url;
    }

    // 文字コードの判定: Content-Type ヘッダ → meta タグ の順で charset を探す
    var html = response.getContentText(); // まず UTF-8 として読む
    var charset = detectCharset_(response.getHeaders(), html);
    if (charset && charset.toLowerCase() !== 'utf-8') {
      // UTF-8 以外なら正しい文字コードで読み直す（Shift_JIS のサイト等の文字化け対策）
      html = response.getContentText(charset);
    }

    return extractTitle_(html) || url;
  } catch (fetchErr) {
    // 取得失敗時は URL そのものをタイトル代わりに使う（仕様のフォールバック）
    return url;
  }
}

/**
 * Content-Type ヘッダまたは HTML の meta タグから charset を検出する。
 * 見つからなければ空文字を返す（= UTF-8 のまま扱う）。
 */
function detectCharset_(headers, html) {
  // ヘッダのキー名は環境により大文字小文字が揺れるため両方見る
  var contentType = String(headers['Content-Type'] || headers['content-type'] || '');
  var m = contentType.match(/charset=["']?([\w-]+)/i);
  if (m) {
    return m[1];
  }
  // <meta charset="..."> または <meta http-equiv="Content-Type" content="...; charset=...">
  m = html.match(/<meta[^>]+charset=["']?([\w-]+)/i);
  if (m) {
    return m[1];
  }
  return '';
}

/**
 * HTML 文字列からページタイトルを抽出する純粋関数。
 * <title> タグを優先し、無ければ og:title を探す。
 * 見つからなければ空文字を返す。
 */
function extractTitle_(html) {
  if (!html) {
    return '';
  }

  // <title> タグ（属性付きの <title data-x="..."> にも対応）
  var m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  var title = m ? m[1] : '';

  // <title> が空なら og:title を探す（content 属性が前後どちらにあっても拾う）
  if (!title.replace(/\s/g, '')) {
    m = html.match(/<meta[^>]+property=["']og:title["'][^>]*content=["']([^"']*)["']/i) ||
        html.match(/<meta[^>]+content=["']([^"']*)["'][^>]*property=["']og:title["']/i);
    title = m ? m[1] : '';
  }

  // HTML実体参照を戻し、改行・連続空白を1つのスペースにまとめる
  return decodeEntities_(title).replace(/\s+/g, ' ').trim();
}

/**
 * 代表的な HTML 実体参照（&amp; &lt; 等）と数値文字参照（&#x27; 等）を
 * 通常の文字に戻す純粋関数。
 */
function decodeEntities_(text) {
  return String(text)
    .replace(/&#x([0-9a-f]+);/gi, function (all, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/&#(\d+);/g, function (all, dec) {
      return String.fromCharCode(parseInt(dec, 10));
    })
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&'); // &amp; は最後に戻す（二重デコード防止）
}

/**
 * メモがタイトルと実質同じ内容かどうかを判定する純粋関数。
 * Android共有時に「タイトル文字列」がそのままメモとして送られてくることが多く、
 * それをそのまま保存すると本文とメモが同じ内容の二重表示になるため、判定して除外する。
 */
function isDuplicateMemo_(memo, title) {
  var normalizedMemo = String(memo || '').trim();
  var normalizedTitle = String(title || '').trim();
  return !!normalizedMemo && normalizedMemo === normalizedTitle;
}

// ---- Drive（過去のGoogleドキュメント参照用） -----------------

/**
 * 親フォルダ直下から名前一致のフォルダを探し、無ければ作成して返す。
 */
function getOrCreateFolder_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * ルートフォルダ「ナレッジ」を取得する（無ければ作成）。
 * 名前検索（getFoldersByName）はマイドライブが大きいと遅いため、
 * 一度見つけたフォルダのIDをスクリプトプロパティに記憶し、
 * 2回目以降はIDで直接開く（カテゴリ一覧・記事一覧APIの応答高速化）。
 * IDのフォルダが削除・ゴミ箱行きになっていた場合は名前検索からやり直す。
 */
function getOrCreateRootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var cachedId = props.getProperty(ROOT_FOLDER_ID_PROPERTY);

  if (cachedId) {
    try {
      var cached = DriveApp.getFolderById(cachedId);
      if (!cached.isTrashed()) {
        return cached;
      }
    } catch (openErr) {
      // IDのフォルダが開けない（削除された等）→ 下の名前検索にフォールバック
    }
  }

  var folder = getOrCreateFolder_(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  props.setProperty(ROOT_FOLDER_ID_PROPERTY, folder.getId());
  return folder;
}

/**
 * 保存する Markdown 本文を組み立てる純粋関数。
 * 新規記事の保存では使わないが、既存Googleドキュメント（action=update）の
 * 本文再構築に引き続き使用する。
 * タイトル・URL・保存日時・カテゴリ（+任意のメモ・本文）を含める。
 * originalUrl が url と異なる場合のみ「共有時のURL」行を追加する
 * （share.google 等の短縮/リダイレクトリンクだった場合の記録用）。
 */
function buildMarkdown_(title, url, savedAt, category, memo, bodyText, originalUrl) {
  var lines = [
    '# ' + title,
    '',
    '- URL: ' + url
  ];
  if (originalUrl && originalUrl !== url) {
    lines.push('- 共有時のURL（短縮/リダイレクト元）: ' + originalUrl);
  }
  lines.push('- 保存日時: ' + savedAt);
  lines.push('- カテゴリ: ' + category);

  if (memo) {
    lines.push('');
    lines.push('## メモ');
    lines.push('');
    lines.push(memo);
  }

  if (bodyText) {
    lines.push('');
    lines.push('## 本文（自動抽出・参考）');
    lines.push('');
    lines.push(bodyText);
  }

  lines.push('');
  return lines.join('\n');
}

// ---- 検索用インデックス（Sheets） -----------------------------

/**
 * 検索用インデックス・スプレッドシートの1枚目のシートを取得する。
 * このスクリプトは「ナレッジ一覧」スプレッドシートにコンテナバインドされている前提のため、
 * getActiveSpreadsheet() で自分自身（バインド先）を直接参照する
 * （openByIdでの外部参照・IDのスクリプトプロパティ保持は不要）。
 * シートが空（初回バインド直後等）ならヘッダー行を書き込んで初期化する。
 * 列追加前から運用しているシート（「タグ」「ID」ヘッダーが無い）は、ヘッダーを補完したうえで
 * migrateRowIds_ で既存行にID（UUID）を後から付与する。
 */
function getOrCreateIndexSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(INDEX_HEADER);
    sheet.setFrozenRows(1);
    return sheet;
  }

  // 不要な列を削除して運用しているシートだと、ID列（H列）自体が存在せず
  // getRange が範囲外エラーになるため、足りない分の列を先に確保する。
  ensureIndexColumns_(sheet);

  // 列追加前から運用しているシートのヘッダー補完（タグ列・ID列）。
  // 空のときだけ書き込むので、既にヘッダーがあるシートでは何もしない。
  if (!sheet.getRange(1, INDEX_COL.TAGS).getValue()) {
    sheet.getRange(1, INDEX_COL.TAGS).setValue(INDEX_HEADER[INDEX_COL.TAGS - 1]);
  }
  if (!sheet.getRange(1, INDEX_COL.ID).getValue()) {
    sheet.getRange(1, INDEX_COL.ID).setValue(INDEX_HEADER[INDEX_COL.ID - 1]);
  }

  migrateRowIds_(sheet);
  return sheet;
}

/**
 * インデックスシートに INDEX_HEADER 分の列数を確保する。
 * 通常のスプレッドシートは初期状態で26列（A〜Z）あるため何もしないが、
 * 余分な列を削除して運用しているシートではID列（H列）が存在せず、
 * getRange(1, INDEX_COL.ID) が範囲外エラーになるため、不足分を追加する。
 */
function ensureIndexColumns_(sheet) {
  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < INDEX_HEADER.length) {
    sheet.insertColumnsAfter(maxColumns, INDEX_HEADER.length - maxColumns);
  }
}

/**
 * ID列が空のデータ行にUUIDを発行する（既存シートの移行用）。
 *
 * 冪等性の担保: 書き込むのは「ID列が空文字（トリム後）の行」だけで、
 * 既に値が入っている行は型を問わず（数値・文字列いずれでも）そのまま温存する。
 * このため何回実行してもID列の内容は初回実行後から変化しない。
 * 空の行が1件も無ければ書き込み自体を行わずに戻る（無駄なsetValuesを避ける）。
 *
 * 既存IDが重複していても修復はしない（「既存の値は触らない」を優先する）。
 * 重複時にどの行が選ばれるかは findRowIndexById_ 側の規約（先頭＝最古を採用）で決まる。
 */
function migrateRowIds_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return; // ヘッダーのみ、またはデータ無し
  }

  var range = sheet.getRange(2, INDEX_COL.ID, lastRow - 1, 1);
  var ids = range.getValues();
  var filled = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] == null ? '' : ids[i][0]).trim() === '') {
      ids[i][0] = Utilities.getUuid();
      filled++;
    }
  }
  if (filled > 0) {
    range.setValues(ids);
  }
}

/**
 * ID列の値が id と一致する行の行番号（1始まり）を返す純粋な検索関数。
 * 見つからなければ -1 を返す。
 *
 * 比較は文字列化・トリムしてから行う（ID列がテキストではなく数値として
 * 保存されているシートでも一致させるため）。
 * ID重複時は先頭（＝最古の行）を採用する。
 */
function findRowIndexById_(sheet, lastRow, id) {
  var needle = String(id == null ? '' : id).trim();
  if (!needle || lastRow < 2) {
    return -1;
  }

  var ids = sheet.getRange(2, INDEX_COL.ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] == null ? '' : ids[i][0]).trim() === needle) {
      return i + 2; // ヘッダ1行 + 0始まりindex の分をずらして実際の行番号にする
    }
  }
  return -1;
}

/**
 * インデックスシートに1行追記する。
 * タイトル列は元記事URLへのHYPERLINK。Driveファイル列は、以前の運用（Googleドキュメント
 * への転記）で作成した記事にのみリンクが入っていたため、新規保存では常に空のままにする。
 * タグ列は複数タグを ", " 区切りで1セルにまとめる（集計・フィルタしやすい形式）。
 * ID列には行を一意に識別するUUIDを必ず発行する（編集・削除時の行特定キー）。
 * 発行したIDは戻り値として返す。
 */
function appendIndexRow_(savedAt, category, title, url, memo, tags) {
  var sheet = getOrCreateIndexSheet_();
  var id = Utilities.getUuid();
  // 自由入力由来の値（タイトル・メモ・カテゴリ・タグ）はセル値としてサニタイズする。
  // GASの appendRow/setValue は先頭が = の文字列を数式として解釈するため、
  // ページタイトルや共有メモに =IMPORTXML(...) 等が入っていると実行されてしまう。
  sheet.appendRow([
    savedAt,
    sanitizeCellText_(category),
    sanitizeCellText_(title),
    url,
    sanitizeCellText_(memo || ''),
    '',
    sanitizeCellText_(tags.join(', ')),
    id
  ]);
  var lastRow = sheet.getLastRow();

  var titleCell = sheet.getRange(lastRow, INDEX_COL.TITLE);
  titleCell.setFormula(
    '=HYPERLINK("' + escapeFormulaString_(url) + '","' + escapeFormulaString_(title) + '")'
  );

  return id;
}

/**
 * スプレッドシートの数式文字列リテラル内に安全に埋め込めるよう、
 * ダブルクォートをエスケープする純粋関数。
 */
function escapeFormulaString_(text) {
  return String(text).replace(/"/g, '""');
}

/**
 * Driveファイル列のHYPERLINK数式から GoogleドキュメントのfileIdを抽出する純粋関数。
 * 例: '=HYPERLINK("https://docs.google.com/document/d/ABC123/edit","開く")' → 'ABC123'
 * 数式が無い・URL形式でない場合は空文字を返す。
 */
function extractFileIdFromFormula_(formula) {
  var m = String(formula || '').match(/document\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

// ---- 保存済み記事の編集（action=update） ----------------------

/**
 * 保存済み記事の内容（タイトル・URL・メモ・タグ）を編集する(action=update)。
 * 行の特定キーは id（ID列のUUID）優先・fileId（GoogleドキュメントのID）フォールバック:
 *   - id指定時: ID列の一致行のSheetsのみ更新する（DocumentApp操作はスキップ。
 *     新規保存の記事はDocを持たないため）。ID重複時は先頭（最古）の一致行を採用する
 *   - fileIdのみ指定時: 従来動作（Sheets行と対応するGoogleドキュメント本文の両方を更新）
 *
 * 行特定は「Sheetsインデックスに載っている記事のみ」に限定する（任意のfileIdで
 * 自分のDrive上の無関係なファイルを触られないよう、一覧に存在する行だけを対象にする）。
 * 保存日時とカテゴリは編集対象外で、Doc再構築時はSheets行の値をそのまま引き継ぐ。
 * ドキュメントの「本文（自動抽出・参考）」節と「共有時のURL」行は保持する。
 * tags: body.tags が配列で指定された場合のみタグ列を置き換える（validateSaveParams_と
 * 同じ方針で、現在のタグ一覧に完全一致するタグのみ許可。未登録タグはエラー）。
 * 未指定（配列でない）の場合はタグ列を変更せず現状を保持する。空配列を指定すると全クリアする。
 */
function handleUpdate_(body) {
  // ステップ1: 入力検証（save と同じ方針: http/https のみ、タイトル必須）
  var id = String(body.id || '').trim();
  var fileId = String(body.fileId || '').trim();
  var title = String(body.title || '').trim();
  var url = String(body.url || '').trim();
  var memo = String(body.memo || '').trim();
  var tagsSpecified = Object.prototype.toString.call(body.tags) === '[object Array]';
  var tags = tagsSpecified ? normalizeTagsInput_(body.tags) : null;

  if (!id && !fileId) {
    throw new Error('編集対象が指定されていません');
  }
  if (!title) {
    throw new Error('タイトルを入力してください');
  }
  if (!/^https?:\/\/\S+$/i.test(url)) {
    throw new Error('URLの形式が不正です: ' + url);
  }
  if (tagsSpecified) {
    var knownTags = getTags_();
    tags.forEach(function (tag) {
      if (knownTags.indexOf(tag) === -1) {
        throw new Error('不明なタグです: ' + tag);
      }
    });
  }

  // ステップ2: Sheetsインデックスから対象行を特定する（id優先・fileIdフォールバック）
  var sheet = getOrCreateIndexSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('編集対象の記事が見つかりません');
  }

  var targetRow = -1;
  if (id) {
    targetRow = findRowIndexById_(sheet, lastRow, id);
  } else {
    var fileFormulas = sheet.getRange(2, INDEX_COL.FILE, lastRow - 1, 1).getFormulas();
    for (var i = 0; i < fileFormulas.length; i++) {
      if (extractFileIdFromFormula_(fileFormulas[i][0]) === fileId) {
        targetRow = i + 2; // ヘッダ1行 + 0始まりindex の分をずらして実際の行番号にする
        break;
      }
    }
  }
  if (targetRow === -1) {
    throw new Error('編集対象の記事が見つかりません');
  }

  // 保存日時・カテゴリは編集しないため、Doc再構築用に現在値をSheetsから引き継ぐ
  var savedAt = String(sheet.getRange(targetRow, INDEX_COL.SAVED_AT).getValue());
  var category = String(sheet.getRange(targetRow, INDEX_COL.CATEGORY).getValue());

  // ステップ3: Sheets行を更新する（URL=平文、メモ=数式インジェクション対策、タイトル=HYPERLINK）
  sheet.getRange(targetRow, INDEX_COL.URL).setValue(sanitizeCellText_(url));
  sheet.getRange(targetRow, INDEX_COL.MEMO).setValue(sanitizeCellText_(memo));
  sheet.getRange(targetRow, INDEX_COL.TITLE).setFormula(
    '=HYPERLINK("' + escapeFormulaString_(url) + '","' + escapeFormulaString_(title) + '")'
  );
  // tags未指定（配列でない）の場合はタグ列に触れず現状を保持する
  if (tagsSpecified) {
    sheet.getRange(targetRow, INDEX_COL.TAGS).setValue(sanitizeCellText_(tags.join(', ')));
  }

  // ステップ4: fileId指定の旧データのみ、Googleドキュメント本文も更新する
  // （自動抽出本文・共有時URLは保持。id指定の記事はDocを持たないためスキップする）
  if (!id) {
    var doc = DocumentApp.openById(fileId);
    var currentText = doc.getBody().getText();
    var newContent = rebuildDocContent_(currentText, savedAt, category, title, url, memo);
    doc.getBody().setText(newContent);
    doc.saveAndClose();
  }

  return { ok: true, id: id, fileId: fileId, title: title, url: url, memo: memo, tags: tagsSpecified ? tags : null };
}

/**
 * 既存ドキュメント本文を、新しいタイトル・URL・メモで組み立て直す純粋関数。
 * 「本文（自動抽出・参考）」節と「共有時のURL」行は元テキストから抽出して引き継ぎ、
 * 保存日時・カテゴリは呼び出し側（Sheets行由来）の値を使う。
 * buildMarkdown_ を再利用するため、save 時と同じ本文フォーマットが保たれる。
 */
function rebuildDocContent_(currentText, savedAt, category, title, url, memo) {
  var originalUrl = extractDocOriginalUrl_(currentText);
  var bodyText = extractDocBodyText_(currentText);
  return buildMarkdown_(title, url, savedAt, category, memo, bodyText, originalUrl);
}

/**
 * ドキュメント本文から「共有時のURL（短縮/リダイレクト元）」の値を抽出する純粋関数。
 * 無ければ空文字を返す。
 */
function extractDocOriginalUrl_(text) {
  var m = String(text || '').match(/^- 共有時のURL（短縮\/リダイレクト元）: (.+)$/m);
  return m ? m[1].trim() : '';
}

/**
 * ドキュメント本文から「本文（自動抽出・参考）」節の中身を抽出する純粋関数。
 * 節見出し以降のテキスト（前後の空白を除く）を返す。節が無ければ空文字。
 */
function extractDocBodyText_(text) {
  var marker = '## 本文（自動抽出・参考）';
  var source = String(text || '');
  var idx = source.indexOf(marker);
  if (idx === -1) {
    return '';
  }
  return source.substring(idx + marker.length).replace(/^\s+/, '').replace(/\s+$/, '');
}

/**
 * セル値として書き込む自由入力テキストを無害化する純粋関数。
 * 先頭が = や + の文字列は数式として解釈されるため、先頭にアポストロフィを付ける
 * （Sheetsのユーザー入力と同じエスケープ。表示・読み取り時にアポストロフィは現れない）。
 */
function sanitizeCellText_(text) {
  var value = String(text);
  if (/^[=+]/.test(value)) {
    return "'" + value;
  }
  return value;
}

// ---- レスポンス ---------------------------------------------

/**
 * オブジェクトを JSON レスポンスとして返す共通関数。
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
