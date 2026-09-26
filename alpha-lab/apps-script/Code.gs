/**
 * Alpha Lab 自動更新スクリプト（Google Apps Script）
 *
 * Claude を使わずに、株価・チャート・マーケット指標・ニュース・経済データ・決算予定・スコア用の財務指標・
 * アナリスト評価を Yahoo Finance などから取得し、Google ドライブの「Alpha Lab データ」フォルダに保存します。
 * Alpha Lab のページは Google Drive コネクタでこのフォルダを読み、ページの更新ボタンはこのフォルダに依頼ファイルを置きます。
 *
 * 使い方（最初の1回だけ）
 *   1. 上の関数の選択欄で「setup」を選び、▶ 実行 を押す
 *   2. 「承認が必要です」→ 権限を確認 → アカウントを選ぶ → 許可
 *   3. 実行ログに「準備ができました」と出れば完了。以後は1分ごとに自動で動きます
 *      （ページのボタンの依頼を処理し、平日の翌朝に株価・スコア・アナリスト評価を自動で更新します）
 *   ※ 新しい版に貼り替えたときも setup をもう一度実行してください（メール送信などの許可を求められます）
 *
 * メール通知: ページの「メール通知」で設定すると、ウォッチリストの銘柄について
 *   価格アラート・大きな値動き・52週高値/安値・格上げ/格下げ・目標株価の変化・決算の前日・EPS予想の修正 をメールで知らせます。
 *   米国市場の取引時間中は15分ごとにも見張ります。
 *
 * 止めたいとき: 関数「stop」を実行します。再開は「setup」をもう一度実行します。
 */

const FOLDER = 'Alpha Lab データ';
const STATUS = 'alpha-lab-status.json';
const REQUEST = 'alpha-lab-request';
const CONFIG = 'alpha-lab-config';
const PAGE = 'https://claude.ai/artifact/5da48G6pbMh6pEwUJkR3Ro';
const SEED = 'https://raw.githubusercontent.com/gukouhehong16-create/baito-tools/alpha-lab-data/quotes.json';
// 自動更新の時刻（日本時間）。dow は曜日（0=日 … 6=土）。米国市場が閉まった後の火〜土の朝に動かす
const SCHEDULE = [
  { mode: 'quotes', dow: [2, 3, 4, 5, 6], at: '06:40' },
  { mode: 'score', dow: [2, 3, 4, 5, 6], at: '06:50' },
  { mode: 'analyst', dow: [2, 3, 4, 5, 6], at: '07:00' },
];
const MODES = ['quotes', 'score', 'analyst'];
const JOBS = { quotes: 'bulk', score: 'score', analyst: 'analyst' };
const FETCH_MS = 250 * 1000; // 1回の実行で取得に使う時間の上限（実行は最長6分）

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36';
const HD = { 'User-Agent': UA, 'Accept': 'application/json, text/plain, */*', 'Accept-Language': 'en-US,en;q=0.9' };
const CNN = Object.assign({}, HD, { 'Origin': 'https://edition.cnn.com', 'Referer': 'https://edition.cnn.com/' });
const Y = 'https://query2.finance.yahoo.com';
const CHUNK = 240000;
const MACRO = ['^GSPC', '^IXIC', '^DJI', '^RUT', '^SOX', '^N225', '^VIX', '^VIX3M', '^MOVE',
  '^IRX', '2YY=F', '^FVX', '^TNX', '^TYX', 'CL=F', 'BZ=F', 'NG=F', 'GC=F', 'SI=F', 'HG=F',
  'JPY=X', 'EURUSD=X', 'DX-Y.NYB', 'BTC-USD', 'ETH-USD'];
const FRED = { DFF: 2, DGS2: 2, DGS10: 2, T10Y2Y: 2, BAMLH0A0HYM2: 2, T10YIE: 2, CPIAUCSL: 4, CPILFESL: 4, PCEPILFE: 4,
  UNRATE: 4, PAYEMS: 4, ICSA: 2, NFCI: 2, MORTGAGE30US: 2, UMCSENT: 4, A191RL1Q225SBEA: 5,
  RECPROUSM156N: 4, WALCL: 3, M2SL: 4 };
const SCOLS = ['roe', 'opm', 'pm', 'gm', 'rg', 'eg', 'pe', 'fpe', 'ps', 'pb', 'ev', 'peg', 'beta', 'dy',
  'ma50', 'ma200', 'hi52', 'lo52', 'roa', 'eps', 'evr', 'gp', 'q'];
const TYPES = [].concat(
  ['TotalRevenue', 'NetIncome', 'DilutedEPS', 'StockholdersEquity', 'TotalAssets'].map(k => 'quarterly' + k),
  ['TotalRevenue', 'NetIncome', 'DilutedEPS'].map(k => 'annual' + k),
  ['TotalRevenue', 'NetIncome', 'OperatingIncome', 'GrossProfit', 'DilutedEPS', 'PeRatio', 'ForwardPeRatio', 'PegRatio',
    'PsRatio', 'PbRatio', 'EnterprisesValueEBITDARatio', 'EnterprisesValueRevenueRatio'].map(k => 'trailing' + k)).join(',');


// ========== 入口 ==========

/** 最初に1回だけ実行する: フォルダ・状態ファイル・1分ごとのトリガーを用意し、接続を確かめてから最初の更新を始める */
function setup() {
  const st = open_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  const d = diagnose_();
  st.diag = d;
  Logger.log('接続テスト: ' + JSON.stringify(d));
  const today = jst_().day;
  const P = props_();
  const sched = P.get('SCHED', {});
  SCHEDULE.forEach(s => { sched[s.mode] = today; });
  P.set('SCHED', sched);
  P.set('QUEUE', MODES.slice());
  P.set('RUN', null);
  st.note = '準備ができました';
  saveStatus_(st);
  Logger.log('準備ができました。フォルダ「' + FOLDER + '」に保存します。続けて最初の更新（株価・チャート）を始めます。');
  const done = tick();
  Logger.log(done ? done + ' の更新が終わりました。残りは、この後1分ごとの自動実行で順に更新されます。'
    : '最初の更新は1分ごとの自動実行で始まります。数分後にフォルダ「' + FOLDER + '」を確認してください。');
}

/** 1分ごとに動く: ページからの依頼と定期更新を受け付け、待ち行列の先頭を1つ実行する（実行した種類を返す） */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return '';
  try {
    const t0 = Date.now(), P = props_(), now = jst_();
    const queue = P.get('QUEUE', []), sched = P.get('SCHED', {});
    let changed = false;
    const add = m => { (m === 'all' ? MODES : [m]).forEach(x => { if (MODES.indexOf(x) >= 0 && queue.indexOf(x) < 0) { queue.push(x); changed = true; } }); };
    const cfgTs = (cfg_() || {}).ts, reqs = takeInbox_(), cfgNew = (cfg_() || {}).ts !== cfgTs;
    reqs.forEach(r => { if (r.mode !== 'testmail') add(r.mode); });
    const wantTest = reqs.some(r => r.mode === 'testmail'), liveDue = watchDue_(P);
    // ページに新しく加わった銘柄（依頼ファイルの add）を対象に加える
    const extra = P.get('EXTRA', []);
    reqs.forEach(r => r.add.forEach(t => { t = String(t).toUpperCase(); if (/^[A-Z0-9][A-Z0-9.\-]{0,11}$/.test(t) && extra.indexOf(t) < 0) extra.push(t); }));
    P.set('EXTRA', extra.slice(0, 400));
    SCHEDULE.forEach(s => {
      if (s.dow.indexOf(now.dow) >= 0 && now.hm >= s.at && sched[s.mode] !== now.day) { sched[s.mode] = now.day; add(s.mode); }
    });
    // 前回の実行が6分の上限で打ち切られた場合: 2回続けて終わらなければ諦める
    const prev = P.get('RUN', null);
    if (prev && queue[0] === prev.mode && prev.tries >= 2) { queue.shift(); changed = true; failLater_(prev.mode, '時間内に終わりませんでした（2回）'); }
    P.set('SCHED', sched);
    P.set('QUEUE', queue);
    const side = st => { if (wantTest) testMail_(st); if (liveDue) watchCheck_(st); st.cfgTs = (cfg_() || {}).ts || null; };
    if (!queue.length) {
      P.set('RUN', null);
      if (changed || reqs.length || liveDue || cfgNew) { const st = open_(); side(st); st.queue = queue; st.running = null; saveStatus_(st); }
      return '';
    }
    const mode = queue[0];
    P.set('RUN', { mode: mode, at: now.at, tries: prev && prev.mode === mode ? (prev.tries || 0) + 1 : 1 });
    const st = open_();
    side(st);
    st.queue = queue.slice();
    st.running = { mode: mode, at: now.at, ts: Date.now() };
    saveStatus_(st);
    const log = [];
    try {
      run_(mode, st, log, t0);
    } catch (e) {
      log.push([JOBS[mode], 'error', '更新できず: ' + msg_(e)]);
    }
    finish_(st, mode, log);
    queue.shift();
    P.set('QUEUE', queue);
    P.set('RUN', null);
    st.queue = queue.slice();
    st.running = null;
    saveStatus_(st);
    return mode;
  } finally {
    lock.releaseLock();
  }
}

/** 今すぐ全部を更新する（手動用） */
function runAll() {
  const P = props_(), q = P.get('QUEUE', []);
  MODES.forEach(m => { if (q.indexOf(m) < 0) q.push(m); });
  P.set('QUEUE', q);
  tick();
}

/** 自動実行を止める */
function stop() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('自動実行を止めました。再開するときは setup を実行してください。');
}

/** 接続テストだけを行う */
function test() {
  Logger.log(JSON.stringify(diagnose_(), null, 1));
}


// ========== 状態・保存 ==========

function props_() {
  const P = PropertiesService.getScriptProperties();
  return {
    get: (k, d) => { const v = P.getProperty(k); if (v == null) return d; try { return JSON.parse(v); } catch (e) { return d; } },
    set: (k, v) => P.setProperty(k, JSON.stringify(v)),
    raw: P,
  };
}

function folder_(P) {
  const id = P.getProperty('FOLDER');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* 作り直す */ } }
  const it = DriveApp.getRootFolder().getFoldersByName(FOLDER);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
  P.setProperty('FOLDER', f.getId());
  return f;
}

function dataFolder_(P, root) {
  const id = P.getProperty('DATA');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* 作り直す */ } }
  const it = root.getFoldersByName('data');
  const f = it.hasNext() ? it.next() : root.createFolder('data');
  P.setProperty('DATA', f.getId());
  return f;
}

/** 状態ファイルを開く（なければ作る）。ページはこのファイルで、どのデータがいつ更新されたかを知る */
function open_() {
  const P = PropertiesService.getScriptProperties(), root = folder_(P), data = dataFolder_(P, root);
  let file = null;
  const id = P.getProperty('STATUS');
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }
  if (!file) { const it = root.getFilesByName(STATUS); file = it.hasNext() ? it.next() : null; }
  let st = null;
  if (file) { try { st = JSON.parse(file.getBlob().getDataAsString('UTF-8')); } catch (e) { st = null; } }
  if (!file) file = root.createFile(STATUS, '{}', 'application/json');
  P.setProperty('STATUS', file.getId());
  st = st && st.app === 'alpha-lab' ? st : { app: 'alpha-lab', v: 1, files: {}, jobs: {}, log: [] };
  st.folderId = root.getId();
  st.dataId = data.getId();
  st.statusId = file.getId();
  st._file = file;
  st._data = data;
  return st;
}

function saveStatus_(st) {
  const now = jst_();
  st.at = now.at;
  st.ts = Date.now();
  const out = {};
  Object.keys(st).forEach(k => { if (k.charAt(0) !== '_') out[k] = st[k]; });
  st._file.setContent(JSON.stringify(out));
}

/** データを gzip → base64 の文字列にして保存する（ファイルIDは変えずに中身だけ差し替える） */
function put_(st, path, obj) {
  const json = JSON.stringify(obj);
  const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
  const md5 = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json, Utilities.Charset.UTF_8));
  const ent = st.files[path];
  let file = null;
  if (ent && ent.id) { try { file = DriveApp.getFileById(ent.id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }
  const name = path.replace('/', '-') + '.json.gz.b64';
  if (!file) { const it = st._data.getFilesByName(name); file = it.hasNext() ? it.next() : null; }
  if (file) file.setContent(b64);
  else file = st._data.createFile(name, b64, 'text/plain');
  st.files[path] = { id: file.getId(), md5: md5, at: jst_().at, n: json.length };
}

/** 前回保存したデータを読む */
function get_(st, path) {
  const ent = st.files[path];
  if (!ent || !ent.id) return null;
  try {
    const b64 = DriveApp.getFileById(ent.id).getBlob().getDataAsString('UTF-8').trim();
    const gz = Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip');
    return JSON.parse(Utilities.ungzip(gz).getDataAsString('UTF-8'));
  } catch (e) {
    return null;
  }
}

/**
 * ページが置いたファイルを読み取って消す。
 *   alpha-lab-request… : 更新の依頼（mode・add）
 *   alpha-lab-config…  : メール通知の設定とウォッチリスト（いちばん新しいものだけを使う）
 */
function takeInbox_() {
  const out = [], cfgs = [];
  const it = DriveApp.searchFiles("(title contains '" + REQUEST + "' or title contains '" + CONFIG + "') and trashed = false");
  while (it.hasNext()) {
    const f = it.next();
    let r = {};
    try { r = JSON.parse(f.getBlob().getDataAsString('UTF-8')) || {}; } catch (e) { r = {}; }
    if (f.getName().indexOf(CONFIG) === 0) cfgs.push([f.getDateCreated().getTime(), r]);
    else out.push({ mode: String(r.mode || 'all'), at: r.at || '', add: Array.isArray(r.add) ? r.add : [] });
    try { f.setTrashed(true); } catch (e) { /* 次回もう一度読む */ }
  }
  if (cfgs.length) {
    cfgs.sort((a, b) => b[0] - a[0]);
    saveCfg_(cfgs[0][1]);
  }
  return out;
}

function finish_(st, mode, log) {
  const now = jst_();
  const runs = (get_(st, 'meta/runs') || {}).items || [];
  log.forEach(l => runs.push({ at: now.at, job: l[0], status: l[1], note: l[2] }));
  runs.sort((a, b) => String(a.at).localeCompare(String(b.at)));
  put_(st, 'meta/runs', { items: runs.slice(-80) });
  const last = log[log.length - 1] || [JOBS[mode], 'error', '結果なし'];
  st.jobs[mode] = { at: now.at, ts: Date.now(), status: last[1], note: last[2] };
  st.log = (st.log || []).concat(log.map(l => [now.at].concat(l))).slice(-30);
  log.forEach(l => Logger.log('[' + l[0] + '] ' + l[1] + ': ' + l[2]));
}

function failLater_(mode, why) {
  const st = open_();
  finish_(st, mode, [[JOBS[mode], 'error', '更新できず: ' + why]]);
  saveStatus_(st);
}


// ========== 共通 ==========

function jst_(ms) {
  const d = new Date((ms || Date.now()) + 9 * 3600 * 1000), iso = d.toISOString();
  return { day: iso.slice(0, 10), hm: iso.slice(11, 16), at: iso.slice(0, 10) + ' ' + iso.slice(11, 16), dow: d.getUTCDay(), date: d };
}

const ymd_ = sec => new Date(sec * 1000).toISOString().slice(0, 10);
const ymdhm_ = sec => { const s = new Date(sec * 1000).toISOString(); return s.slice(0, 10) + ' ' + s.slice(11, 16); };
const hex_ = bytes => bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
const ysym_ = t => t.replace(/\./g, '-');
const at_ = (a, i) => (a && i < a.length && a[i] != null ? a[i] : null);
const rnd_ = (x, n) => { const k = Math.pow(10, n); return Math.round(x * k) / k; };
const sig_ = x => Number(Number(x).toPrecision(5));
const r4_ = x => rnd_(x, 4).toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
const msg_ = e => String((e && e.message) || e).slice(0, 60);
const days_ = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 86400000);

function num_(x) {
  if (x && typeof x === 'object') x = x.raw;
  if (typeof x !== 'number' || !isFinite(x)) return null;
  return Math.abs(x) >= 1e5 ? Math.round(x) : rnd_(x, 4);
}

function utf8len_(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); n += c < 0x80 ? 1 : c < 0x800 ? 2 : (c >= 0xd800 && c < 0xdc00) ? 2 : 3; }
  return n;
}

/**
 * まとめて並列に取得する。429（混雑）や 5xx は少し待って最大3回まで取り直す。
 * 接続できないもの（時間切れなど）は1回で諦める（1件ごとに最大1分待たされ、6分の実行上限を超えるため）。
 * 戻り値は URL と同じ順の [{code, text}]（取れなかったものは code が数値以外になる）
 */
function many_(urls, headers, deadline, batch) {
  const n = batch || 20, out = urls.map(() => null);
  let todo = urls.map((u, i) => i);
  for (let round = 0; round < 3 && todo.length; round++) {
    if (round) Utilities.sleep(2500 * round);
    const again = [];
    for (let s = 0; s < todo.length; s += n) {
      if (deadline && Date.now() > deadline) { todo.slice(s).forEach(i => { out[i] = out[i] || { code: 'time' }; }); break; }
      const part = todo.slice(s, s + n);
      let rs;
      try {
        rs = UrlFetchApp.fetchAll(part.map(i => ({ url: urls[i], headers: headers || HD, muteHttpExceptions: true })));
      } catch (e) {
        // どれかに接続できなかった: 1件ずつ取り直す。2件続けて接続できなければ残りは諦める
        let bad = part.length === 1 ? 2 : 0;
        rs = part.map(i => {
          if (bad >= 2 || (deadline && Date.now() > deadline)) return e;
          try { const r = UrlFetchApp.fetch(urls[i], { headers: headers || HD, muteHttpExceptions: true }); bad = 0; return r; } catch (e2) { bad++; return e2; }
        });
      }
      let busy = false;
      part.forEach((i, j) => {
        const r = rs[j];
        if (!r || typeof r.getResponseCode !== 'function') { out[i] = { code: 'net ' + msg_(r) }; return; }
        const code = r.getResponseCode();
        if (code === 200) { out[i] = { code: 200, text: r.getContentText() }; return; }
        out[i] = { code: code };
        if (code === 429 || code >= 500) { again.push(i); busy = busy || code === 429; }
      });
      if (busy) Utilities.sleep(3000);
    }
    todo = again;
  }
  return out;
}

function one_(url, headers) {
  const r = many_([url], headers)[0];
  if (!r || r.code !== 200) throw new Error(r && typeof r.code === 'number' ? 'http' + r.code : String((r && r.code) || 'no_response'));
  return r.text;
}

const json_ = r => { if (!r || r.code !== 200) throw new Error(r && typeof r.code === 'number' ? 'http' + r.code : String((r && r.code) || 'no_response')); return JSON.parse(r.text); };

/** 各 URL を取得して f で変換する。失敗は {err} で返す */
function mapFetch_(urls, f, headers, deadline, batch) {
  return many_(urls, headers, deadline, batch).map(r => { try { return { v: f(json_(r)) }; } catch (e) { return { err: msg_(e) }; } });
}

/** Yahoo Finance の認証用トークン（クッキーを受け取ってから取得する） */
function crumb_() {
  const pre = ['https://fc.yahoo.com/', 'https://finance.yahoo.com/quote/AAPL/'];
  for (let p = 0; p < pre.length; p++) {
    let cookie = '';
    try {
      const r = UrlFetchApp.fetch(pre[p], { headers: HD, muteHttpExceptions: true, followRedirects: false });
      const h = r.getAllHeaders();
      let sc = h['Set-Cookie'] || h['set-cookie'] || [];
      if (!Array.isArray(sc)) sc = [sc];
      cookie = sc.map(s => String(s).split(';')[0]).filter(Boolean).join('; ');
    } catch (e) { cookie = ''; }
    if (!cookie) continue;
    const hd = Object.assign({}, HD, { Cookie: cookie });
    for (const host of ['query2', 'query1']) {
      try {
        const r = UrlFetchApp.fetch('https://' + host + '.finance.yahoo.com/v1/test/getcrumb', { headers: hd, muteHttpExceptions: true });
        const c = r.getContentText().trim();
        if (r.getResponseCode() === 200 && c && c.indexOf('<') < 0 && c.length < 40) return { crumb: c, headers: hd };
      } catch (e) { /* 次を試す */ }
    }
  }
  return null;
}

function chunks_(d, prefix) {
  const out = [];
  let cur = {}, size = 0, has = false;
  Object.keys(d).sort().forEach(t => {
    const s = utf8len_(JSON.stringify(d[t])) + t.length + 6;
    if (has && size + s > CHUNK) { out.push(cur); cur = {}; size = 0; has = false; }
    cur[t] = d[t]; size += s; has = true;
  });
  if (has) out.push(cur);
  return out.map((c, i) => [prefix + ('0' + i).slice(-2), c]);
}

function tickers_(st) {
  let q = get_(st, 'meta/quotes');
  if (!q || !q.rows) {
    try { q = JSON.parse(one_(SEED)); } catch (e) { q = null; }
  }
  if (!q || !q.rows) throw new Error('銘柄の一覧を読めません（初回はGitHubから読み込みます）');
  const all = {};
  Object.keys(q.rows).concat(props_().get('EXTRA', [])).forEach(t => { all[t] = 1; });
  return { q: q, tickers: Object.keys(all).sort() };
}

function run_(mode, st, log, t0) {
  const deadline = t0 + FETCH_MS;
  const u = tickers_(st);
  if (mode === 'quotes') runQuotes_(st, u.tickers, u.q, log, deadline, t0);
  else if (mode === 'score') runScore_(st, u.tickers, crumb_(), log, deadline, t0);
  else if (mode === 'analyst') runAnalyst_(st, u.tickers, crumb_(), log, deadline, t0);
}


// ========== 株価・チャート・マーケット指標・ニュース・経済データ・予定 ==========

function parseChart_(j) {
  const r = ((j.chart || {}).result || [null])[0];
  if (!r) throw new Error('no_data');
  const m = r.meta, q = r.indicators.quote[0], off = m.gmtoffset != null ? m.gmtoffset : -14400, ts = r.timestamp || [];
  const bars = {};
  for (let i = 0; i < ts.length; i++) {
    const o = at_(q.open, i), h = at_(q.high, i), l = at_(q.low, i), c = at_(q.close, i), v = at_(q.volume, i);
    if (o == null || h == null || l == null || c == null || c <= 0) continue;
    bars[ymd_(ts[i] + off)] = [o, Math.max(h, o, c), Math.min(l, o, c), c, v || 0];
  }
  const days = Object.keys(bars).sort();
  if (!days.length) throw new Error('no_bars');
  const now = Date.now() / 1000, today = ymd_(now + off);
  const end = ((m.currentTradingPeriod || {}).regular || {}).end || 0;
  const live = days[days.length - 1] === today && now < end; // 取引中の当日足は未確定なのでチャートに入れない
  const px = m.regularMarketPrice || bars[days[days.length - 1]][3];
  const pt = ymdhm_((m.regularMarketTime || now) + off);
  if (live) delete bars[today];
  return { bars: bars, p: px, pt: pt, live: live };
}

function parseMacro_(j) {
  const r = ((j.chart || {}).result || [null])[0];
  if (!r) throw new Error('no_data');
  const m = r.meta, q = (r.indicators.quote || [{}])[0] || {}, off = m.gmtoffset || 0, t = [], c = [], ts = r.timestamp || [];
  for (let i = 0; i < ts.length; i++) {
    const v = at_(q.close, i);
    if (v == null || v <= 0) continue;
    const d = Math.floor((ts[i] + off) / 86400);
    if (t.length && t[t.length - 1] === d) c[c.length - 1] = v;
    else { t.push(d); c.push(v); }
  }
  if (c.length < 5) throw new Error('no_bars');
  return { t: t, c: c.map(sig_), p: sig_(m.regularMarketPrice || c[c.length - 1]), tm: m.regularMarketTime != null ? m.regularMarketTime : null };
}

function parseNews_(t) {
  const s = ysym_(t);
  return j => {
    const out = [];
    const ns = j.news || [];
    for (let i = 0; i < ns.length; i++) {
      const n = ns[i], rel = n.relatedTickers || [];
      if (n.title && (!rel.length || rel.indexOf(s) >= 0 || rel.indexOf(t) >= 0)) {
        out.push([n.title.slice(0, 170), (n.publisher || '').slice(0, 40), n.providerPublishTime || 0, (n.link || '').slice(0, 300)]);
      }
      if (out.length === 5) break;
    }
    return out;
  };
}

function parseFred_(txt) {
  const t = [], c = [];
  txt.trim().split(/\r?\n/).slice(1).forEach(line => {
    const k = line.indexOf(','), d = line.slice(0, k).trim(), v = Number(line.slice(k + 1));
    if (k < 0 || !/^\d{4}-\d{2}-\d{2}$/.test(d) || line.slice(k + 1).trim() === '' || !isFinite(v)) return;
    t.push(Math.floor(Date.parse(d + 'T00:00:00Z') / 86400000));
    c.push(sig_(v));
  });
  if (!c.length) throw new Error('no_data');
  return { t: t, c: c };
}

function fredUrl_(sid, years) {
  const start = jst_(Date.now() - 365 * years * 86400000).day;
  return 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=' + sid + '&cosd=' + start;
}

function parseFng_(j) {
  const f = j.fear_and_greed;
  if (!f) throw new Error('no_data');
  const r1 = v => (typeof v === 'number' ? rnd_(v, 1) : null);
  const keys = ['market_momentum_sp500', 'stock_price_strength', 'stock_price_breadth', 'put_call_options',
    'market_volatility_vix', 'junk_bond_demand', 'safe_haven_demand'];
  const c = {};
  keys.forEach(k => { if (j[k]) c[k] = [r1(j[k].score), j[k].rating != null ? j[k].rating : '']; });
  return {
    s: r1(f.score), r: f.rating != null ? f.rating : '', pc: r1(f.previous_close), w1: r1(f.previous_1_week),
    m1: r1(f.previous_1_month), y1: r1(f.previous_1_year), ts: f.timestamp != null ? f.timestamp : '',
    h: ((j.fear_and_greed_historical || {}).data || []).map(p => [Math.floor(p.x / 86400000), r1(p.y)]),
    c: c,
  };
}

function runQuotes_(st, tickers, q, log, deadline, t0) {
  const now = jst_();
  const AT = now.at;
  const charts = mapFetch_(tickers.map(t => Y + '/v8/finance/chart/' + ysym_(t) + '?range=1y&interval=1d'), parseChart_, HD, deadline, 25);
  const res = {}, err = {};
  tickers.forEach((t, i) => { if (charts[i].v) res[t] = charts[i].v; else err[t] = 'chart_missing'; });
  const n = Object.keys(res).length;
  if (n < tickers.length * 0.8) {
    const why = Object.keys(err).length ? charts[tickers.indexOf(Object.keys(err)[0])].err : '?';
    log.push(['bulk', 'error', '株価・チャートを更新できず: Yahoo Finance から取得できたのが ' + n + '/' + tickers.length + '社のみ（' + why + '）']);
    return;
  }
  // チャート（1年分）
  const rows = {};
  Object.keys(res).forEach(t => {
    const b = res[t].bars;
    rows[t] = Object.keys(b).sort().map(d => d + ',' + b[d].slice(0, 4).map(r4_).join(',') + ',' + Math.trunc(b[d][4])).join(';');
  });
  const cmap = {};
  const cch = chunks_(rows, 'c');
  cch.forEach(c => { put_(st, 'bars/' + c[0], { at: AT, d: c[1] }); Object.keys(c[1]).forEach(t => { cmap[t] = c[0]; }); });
  // 株価: 時価総額は株数据え置きで株価に比例、PER は利益据え置きで株価に比例
  const oldRows = q.rows || {}, oldPe = q.pe || {}, newRows = {}, newPe = {};
  let live = 0, asof = '', lastBar = '';
  Object.keys(res).forEach(t => {
    const r = res[t], p = r.p, old = oldRows[t] || [null, null, null, null];
    const days = Object.keys(r.bars).sort(), y0 = days.filter(d => d < r.pt.slice(0, 4) + '-01-01');
    const k = old[1] ? p / old[1] : null;
    const mc = k && old[0] ? old[0] * k : old[0];
    newRows[t] = [mc ? rnd_(mc, mc < 100 ? 2 : 1) : mc, rnd_(p, 2), old[2], y0.length ? rnd_(p / r.bars[y0[y0.length - 1]][3] - 1, 4) : old[3]];
    if (k && oldPe[t]) newPe[t] = rnd_(oldPe[t] * k, 1);
    if (r.live) live++;
    if (r.pt.slice(0, 10) > asof) asof = r.pt.slice(0, 10);
    if (days.length && days[days.length - 1] > lastBar) lastBar = days[days.length - 1];
  });
  Object.keys(oldRows).forEach(t => { if (!newRows[t]) newRows[t] = oldRows[t]; }); // 取れなかった銘柄は前回の値を残す
  put_(st, 'meta/quotes', { rows: newRows, pe: Object.assign({}, oldPe, newPe), asOf: asof, updatedAt: AT });
  // マーケット指標
  const mres = mapFetch_(MACRO.map(s => Y + '/v8/finance/chart/' + encodeURIComponent(s) + '?range=1y&interval=1d'), parseMacro_, HD, deadline, 25);
  const mac = {}, merr = [];
  MACRO.forEach((s, i) => { if (mres[i].v) mac[s] = mres[i].v; else merr.push(s); });
  if (Object.keys(mac).length) put_(st, 'meta/macro', { at: AT, items: mac, err: merr });
  // ニュース
  const nres = many_(tickers.map(t => Y + '/v1/finance/search?q=' + encodeURIComponent(ysym_(t)) + '&quotesCount=0&newsCount=8&enableFuzzyQuery=false'),
    HD, deadline, 25);
  const nw = {};
  tickers.forEach((t, i) => { try { const v = parseNews_(t)(json_(nres[i])); if (v.length) nw[t] = v; } catch (e) { /* この銘柄はなし */ } });
  const nmap = {};
  chunks_(nw, 'n').forEach(c => { put_(st, 'news/' + c[0], { at: AT, d: c[1] }); Object.keys(c[1]).forEach(t => { nmap[t] = c[0]; }); });
  // 経済データ・Fear & Greed・話題の銘柄
  const diag = {}, oldEcon = get_(st, 'meta/econ') || {};
  const econ = { at: AT, fred: {}, fng: null, trend: [], diag: diag };
  try {
    const tq = (JSON.parse(one_(Y + '/v1/finance/trending/US?count=25')).finance || {}).result || [{}];
    econ.trend = ((tq[0] || {}).quotes || []).filter(x => x.symbol).map(x => x.symbol);
  } catch (e) { diag.trend = msg_(e); }
  try {
    parseFred_(one_(fredUrl_('DFF', 1), HD));
    const ids = Object.keys(FRED), fr = many_(ids.map(s => fredUrl_(s, FRED[s])), HD, deadline, 10);
    ids.forEach((s, i) => { try { if (fr[i] && fr[i].code === 200) econ.fred[s] = parseFred_(fr[i].text); } catch (e) { /* この系列はなし */ } });
  } catch (e) { diag.fred = msg_(e); }
  try { econ.fng = parseFng_(JSON.parse(one_('https://production.dataviz.cnn.io/index/fearandgreed/graphdata/' + jst_(Date.now() - 370 * 86400000).day, CNN))); }
  catch (e) { diag.fng = msg_(e); }
  const gotFred = Object.keys(econ.fred).length > 0;
  if (gotFred || econ.fng) econ.fat = AT;
  if (!gotFred) econ.fred = oldEcon.fred || {};
  econ.fng = econ.fng || oldEcon.fng || null;
  if (!econ.fat) econ.fat = oldEcon.fat || null;
  put_(st, 'meta/econ', econ);
  // ウォッチリストの見張り（終値ベース）と、取引時間中の見張りに使う52週高値・安値
  try { alertsQuotes_(st, res); } catch (e) { st.mailErr = '通知の判定に失敗: ' + msg_(e); }
  const sec = Math.round((Date.now() - t0) / 1000), ne = Object.keys(err).length, dk = Object.keys(diag);
  const note = '一斉更新: 株価' + n + '社（' + asof + (live ? '・取引中の値を含む' : '') + '）、チャート日足1年分' + Object.keys(rows).length + '社、'
    + 'マーケット指標' + Object.keys(mac).length + '種、ニュース' + Object.keys(nw).length + '社、経済データ' + Object.keys(econ.fred).length + '系列'
    + '（Google Apps Script、' + sec + '秒）' + (ne ? '、取得できず ' + ne + '社' : '') + (dk.length ? '、接続できず ' + dk.join(', ') : '');
  put_(st, 'meta/bulk', { at: AT, day: now.day, asof: asof, last: lastBar, src: 'Yahoo Finance', n: n, live: live, err: err, note: note,
    map: cmap, chunks: cch.length, nmap: nmap, via: 'gas' });
  log.push(['bulk', ne || dk.length ? 'partial' : 'ok', note]);
}


// ========== スコア用の財務指標 ==========

function parseTs_(j) {
  const out = {};
  ((j.timeseries || {}).result || []).forEach(r => {
    const k = ((r.meta || {}).type || [null])[0];
    const pts = (k && r[k] || []).filter(x => x && typeof x === 'object' && x.asOfDate && num_(x.reportedValue) != null)
      .map(x => [x.asOfDate, num_(x.reportedValue)]).sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]));
    if (k && pts.length) out[k] = pts;
  });
  if (!Object.keys(out).some(k => k.indexOf('quarterly') === 0 || k.indexOf('annual') === 0)) throw new Error('no_financials');
  return out;
}

function parseHist_(j) {
  const r = ((j.chart || {}).result || [null])[0];
  if (!r) throw new Error('no_chart');
  const q = ((r.indicators || {}).quote || [{}])[0] || {}, ts = r.timestamp || [], cl = {}, hi = [], lo = [];
  for (let i = 0; i < ts.length; i++) {
    const c = at_(q.close, i), h = at_(q.high, i), l = at_(q.low, i);
    if (c && c > 0) { cl[Math.floor(ts[i] / 86400)] = c; hi.push(h || c); lo.push(l || c); }
  }
  const cut = Date.now() / 1000 - 365 * 86400, dvs = (r.events || {}).dividends || {};
  let dv = 0;
  Object.keys(dvs).forEach(k => { const d = dvs[k]; if ((d.date || 0) > cut) dv += d.amount || 0; });
  return { cl: cl, hi: hi.length ? Math.max.apply(null, hi) : null, lo: lo.length ? Math.min.apply(null, lo) : null, dv: dv };
}

function rets_(cl) {
  const ks = Object.keys(cl).map(Number).sort((a, b) => a - b), out = {};
  for (let i = 1; i < ks.length; i++) out[ks[i]] = cl[ks[i]] / cl[ks[i - 1]] - 1;
  return out;
}

function beta_(cl, mk) {
  const r = rets_(cl), ds = Object.keys(r).map(Number).sort((a, b) => a - b).filter(d => d in mk);
  if (ds.length < 60) return null;
  const x = ds.map(d => mk[d]), y = ds.map(d => r[d]);
  const mx = x.reduce((a, b) => a + b, 0) / x.length, my = y.reduce((a, b) => a + b, 0) / y.length;
  let vx = 0, cv = 0;
  for (let i = 0; i < x.length; i++) { vx += (x[i] - mx) * (x[i] - mx); cv += (x[i] - mx) * (y[i] - my); }
  return vx ? cv / vx : null;
}

function scoreRow_(ts, ch, mk) {
  const last = k => (ts[k] && ts[k].length ? ts[k][ts[k].length - 1][1] : null);
  const ttm = k => {
    let v = last('trailing' + k);
    const q = ts['quarterly' + k] || [];
    if (v == null && q.length >= 4 && days_(q[q.length - 4][0], q[q.length - 1][0]) < 300) v = q.slice(-4).reduce((a, x) => a + x[1], 0);
    return v;
  };
  const yoy = k => {
    const q = ts['quarterly' + k] || [];
    for (let i = 0; i < q.length - 1; i++) {
      const dd = days_(q[i][0], q[q.length - 1][0]);
      if (dd >= 330 && dd <= 400) return q[i][1] > 0 ? q[q.length - 1][1] / q[i][1] - 1 : null;
    }
    const a = ts['annual' + k] || [];
    return a.length >= 2 && a[a.length - 2][1] > 0 ? a[a.length - 1][1] / a[a.length - 2][1] - 1 : null;
  };
  const div = (a, b) => (a != null && b ? a / b : null);
  let rev = ttm('TotalRevenue');
  const ni = ttm('NetIncome'), eq = last('quarterlyStockholdersEquity'), ta = last('quarterlyTotalAssets');
  rev = rev && rev > 0 ? rev : null;
  const eg = yoy('DilutedEPS');
  let q = null;
  Object.keys(ts).forEach(k => { if (k.indexOf('quarterly') === 0) { const d = ts[k][ts[k].length - 1][0]; if (q == null || d > q) q = d; } });
  const row = {
    roe: eq && eq > 0 ? div(ni, eq) : null, opm: div(ttm('OperatingIncome'), rev), pm: div(ni, rev),
    gm: div(ttm('GrossProfit'), rev), rg: yoy('TotalRevenue'), eg: eg != null ? eg : yoy('NetIncome'),
    pe: last('trailingPeRatio'), fpe: last('trailingForwardPeRatio'), ps: last('trailingPsRatio'),
    pb: last('trailingPbRatio'), ev: last('trailingEnterprisesValueEBITDARatio'), peg: last('trailingPegRatio'),
    roa: div(ni, ta), eps: ttm('DilutedEPS'), evr: last('trailingEnterprisesValueRevenueRatio'),
    gp: ttm('GrossProfit'), q: q,
  };
  if (ch) {
    const c = Object.keys(ch.cl).map(Number).sort((a, b) => a - b).map(d => ch.cl[d]);
    const avg = a => a.reduce((x, y) => x + y, 0) / a.length;
    row.beta = mk ? beta_(ch.cl, mk) : null;
    row.dy = c.length ? ch.dv / c[c.length - 1] : null;
    row.ma50 = c.length >= 50 ? avg(c.slice(-50)) : null;
    row.ma200 = c.length >= 200 ? avg(c.slice(-200)) : null;
    row.hi52 = ch.hi;
    row.lo52 = ch.lo;
  }
  return SCOLS.map(k => (k === 'q' ? (row[k] == null ? null : row[k]) : num_(row[k])));
}

function runScore_(st, tickers, cr, log, deadline, t0) {
  const now = jst_(), hd = cr ? cr.headers : HD, p1 = Math.floor(Date.now() / 1000) - 3 * 366 * 86400, p2 = Math.floor(Date.now() / 1000) + 86400;
  let mk = null;
  try { mk = rets_(parseHist_(JSON.parse(one_(Y + '/v8/finance/chart/SPY?range=1y&interval=1d&events=div'))).cl); } catch (e) { mk = null; }
  if (mk && !Object.keys(mk).length) mk = null;
  const tsr = mapFetch_(tickers.map(t => {
    const s = ysym_(t);
    return Y + '/ws/fundamentals-timeseries/v1/finance/timeseries/' + s + '?symbol=' + s + '&type=' + TYPES + '&period1=' + p1 + '&period2=' + p2
      + (cr ? '&crumb=' + encodeURIComponent(cr.crumb) : '');
  }), parseTs_, hd, deadline, 20);
  const hist = mapFetch_(tickers.map(t => Y + '/v8/finance/chart/' + ysym_(t) + '?range=1y&interval=1d&events=div'), parseHist_, HD, deadline, 25);
  const rows = {}, err = {};
  tickers.forEach((t, i) => {
    if (!tsr[i].v) { err[t] = tsr[i].err; return; }
    try { rows[t] = scoreRow_(tsr[i].v, hist[i].v || null, mk); } catch (e) { err[t] = msg_(e).slice(0, 40); }
  });
  const n = Object.keys(rows).length;
  if (n < tickers.length * 0.6) {
    log.push(['score', 'error', 'スコア用の財務指標を更新できず: 取得できたのが ' + n + '/' + tickers.length + '社のみ（主な原因 ' + top_(err) + '）']);
    return;
  }
  const prev = (get_(st, 'meta/fund') || {}).rows || {}, all = {};
  Object.keys(err).forEach(t => { if (prev[t]) all[t] = prev[t]; });
  Object.keys(rows).forEach(t => { all[t] = rows[t]; });
  put_(st, 'meta/fund', { at: now.at, src: 'Yahoo Finance', n: n, err: err, cols: SCOLS, rows: all });
  const ne = Object.keys(err).length;
  log.push(['score', ne ? 'partial' : 'ok', 'スコア用の財務指標 ' + n + '社（Google Apps Script、' + Math.round((Date.now() - t0) / 1000) + '秒）' + (ne ? '、取得できず ' + ne + '社' : '')]);
}

function top_(err) {
  const c = {};
  Object.keys(err).forEach(t => { c[err[t]] = (c[err[t]] || 0) + 1; });
  const ks = Object.keys(c).sort((a, b) => c[b] - c[a]);
  return ks.length ? ks[0] : '?';
}


// ========== アナリスト評価 ==========

function parseAnalyst_(j) {
  const r = ((j.quoteSummary || {}).result || [null])[0];
  if (!r) throw new Error('no_result');
  const f = r.financialData || {}, tr = {};
  ((r.recommendationTrend || {}).trend || []).forEach(x => { tr[x.period] = ['strongBuy', 'buy', 'hold', 'sell', 'strongSell'].map(k => num_(x[k]) || 0); });
  const hist = ((r.upgradeDowngradeHistory || {}).history || []).slice().sort((a, b) => (b.epochGradeDate || 0) - (a.epochGradeDate || 0)).slice(0, 3);
  const ud = hist.map(h => [ymd_(h.epochGradeDate || 0), h.firm || '', h.action || '', h.fromGrade || '', h.toGrade || '', num_(h.currentPriceTarget)]);
  const row = {
    m: num_(f.recommendationMean), k: f.recommendationKey || '', n: num_(f.numberOfAnalystOpinions), p: num_(f.currentPrice),
    t: [num_(f.targetMeanPrice), num_(f.targetHighPrice), num_(f.targetLowPrice)], tr: tr['0m'] || null, tr1: tr['-1m'] || null, ud: ud,
  };
  if (!(row.m || (row.tr || []).some(x => x) || ud.length)) throw new Error('no_coverage');
  return row;
}

function parseInsights_(j) {
  const r = (j.finance || {}).result || {}, rec = r.recommendation || {};
  const rep = (r.reports || []).filter(x => x && typeof x === 'object' && x.investmentRating)
    .sort((a, b) => String(b.reportDate || '').localeCompare(String(a.reportDate || '')));
  const val = (r.instrumentInfo || {}).valuation || {};
  const row = {
    pr: String(rec.rating || '').toUpperCase(), pv: rec.provider || '', tp: num_(rec.targetPrice),
    rr: rep.length ? [String(rep[0].reportDate || '').slice(0, 10), rep[0].provider || '', rep[0].investmentRating || '', num_(rep[0].targetPrice), rep[0].targetPriceStatus || ''] : null,
    va: val.description || '',
  };
  if (!(row.pr || row.rr)) throw new Error('no_coverage');
  return row;
}

/** コンセンサス（要トークン）に、調査会社（Argus など）の評価を重ねる */
function runAnalyst_(st, tickers, cr, log, deadline, t0) {
  const now = jst_();
  // 1社につき1回の問い合わせで、推奨・目標株価・格付け変更と、決算日・EPS予想の推移・過去の決算サプライズをまとめて取る
  const raw = cr ? many_(tickers.map(t => Y + '/v10/finance/quoteSummary/' + ysym_(t)
    + '?modules=financialData,recommendationTrend,upgradeDowngradeHistory,calendarEvents,earningsTrend,earningsHistory&crumb=' + encodeURIComponent(cr.crumb)),
  cr.headers, deadline, 20) : null;
  const sum = [], earn = {};
  tickers.forEach((t, i) => {
    if (!raw) { sum.push(null); return; }
    let j;
    try { j = json_(raw[i]); } catch (e) { sum.push({ err: msg_(e) }); return; }
    try { sum.push({ v: parseAnalyst_(j) }); } catch (e) { sum.push({ err: msg_(e) }); }
    try { const e = parseEarn_(j); if (e) earn[t] = e; } catch (e) { /* 決算データなし */ }
  });
  const ins = mapFetch_(tickers.map(t => Y + '/ws/insights/v2/finance/insights?symbol=' + ysym_(t) + '&lang=en-US&region=US'), parseInsights_, HD, deadline, 20);
  const rows = {}, err = {};
  tickers.forEach((t, i) => {
    const s = sum[i];
    let row = s && s.v ? s.v : null;
    const errs = [];
    if (s && !s.v) errs.push(String(s.err).slice(0, 30));
    if (ins[i].v) {
      if (!row) row = { m: null, k: '', n: null, p: null, t: [ins[i].v.tp, null, null], tr: null, tr1: null, ud: [] };
      Object.keys(ins[i].v).forEach(k => { if (k !== 'tp') row[k] = ins[i].v[k]; });
    } else errs.push(String(ins[i].err).slice(0, 30));
    if (row) rows[t] = row;
    else err[t] = errs[0] || 'no_coverage';
  });
  const n = Object.keys(rows).length;
  if (n < tickers.length * 0.3) {
    log.push(['analyst', 'error', 'アナリスト評価を更新できず: 取得できたのが ' + n + '/' + tickers.length + '社のみ（トークン' + (cr ? 'あり' : 'なし') + '、主な原因 ' + top_(err) + '）']);
    return;
  }
  const prev = (get_(st, 'meta/analyst') || {}).rows || {}, all = {};
  Object.keys(err).forEach(t => { if (prev[t]) all[t] = prev[t]; });
  Object.keys(rows).forEach(t => { all[t] = rows[t]; });
  const src = cr ? 'Yahoo Finance' : 'Yahoo Finance（調査会社の評価）', ne = Object.keys(err).length;
  put_(st, 'meta/analyst', { at: now.at, src: src, n: n, err: err, rows: all });
  // 決算日・EPS予想の修正・決算サプライズと、今後2週間の決算・経済指標の予定
  const nE = Object.keys(earn).length;
  let allE = (get_(st, 'meta/earn') || {}).rows || {}, calNote = '';
  if (nE >= tickers.length * 0.3) {
    const keep = {};
    tickers.forEach(t => { if (!earn[t] && allE[t]) keep[t] = allE[t]; });
    allE = Object.assign(keep, earn);
    put_(st, 'meta/earn', { at: now.at, src: 'Yahoo Finance', cols: ECOLS, rows: allE });
    const cal = { at: now.at, src: 'Yahoo Finance', earn: calEarn_(allE), econ: [] };
    try { cal.econ = econEvents_(cr); } catch (e) { cal.econErr = msg_(e); cal.econ = (get_(st, 'meta/cal') || {}).econ || []; }
    put_(st, 'meta/cal', cal);
    calNote = '、決算と予想修正 ' + nE + '社・今後2週間の決算 ' + cal.earn.length + '件・経済指標 ' + cal.econ.length + '件';
  }
  try { alertsAnalyst_(st, prev, all, allE); } catch (e) { st.mailErr = '通知の判定に失敗: ' + msg_(e); }
  log.push(['analyst', ne ? 'partial' : 'ok', 'アナリスト評価 ' + n + '社（' + src + '、Google Apps Script、' + Math.round((Date.now() - t0) / 1000) + '秒）'
    + (ne ? '、取得できず ' + ne + '社' : '') + (cr ? '' : '、コンセンサスは取得できず') + calNote]);
}


// ========== 決算と予想修正 ==========

// 決算データの列: 次回決算日・時間帯（BMO=寄付前, AMC=引け後）・EPS予想・予想人数・前年同期EPS・対象四半期の末日・
// 今期EPS予想の7/30/90日の変化率・来期の30日の変化率・30日の上方/下方修正の件数・過去4四半期のサプライズ率と四半期・日付が推定か・売上予想（10億ドル）
const ECOLS = ['nd', 'tm', 'est', 'ne', 'ya', 'fq', 'r7', 'r30', 'r90', 'n30', 'u30', 'd30', 'sp', 'sq', 'ds', 'rev'];

function parseEarn_(j) {
  const r = ((j.quoteSummary || {}).result || [null])[0];
  if (!r) return null;
  const ce = (r.calendarEvents || {}).earnings || {}, ed = (ce.earningsDate || [])[0], tr = {};
  ((r.earningsTrend || {}).trend || []).forEach(x => { if (x && x.period) tr[x.period] = x; });
  const q0 = tr['0q'] || {}, y0 = tr['0y'] || {}, y1 = tr['+1y'] || {}, ee = q0.earningsEstimate || {};
  const rv = (x, k) => {
    const e = x.epsTrend || {}, c = num_(e.current), a = num_(e[k]);
    return c != null && a != null && Math.abs(a) >= 0.01 ? rnd_((c - a) / Math.abs(a), 4) : null;
  };
  const revs = (y0.epsRevisions && Object.keys(y0.epsRevisions).length ? y0 : q0).epsRevisions || {};
  const hist = ((r.earningsHistory || {}).history || []).filter(h => h && h.quarter && num_(h.quarter) != null)
    .sort((a, b) => num_(a.quarter) - num_(b.quarter)).slice(-4);
  let nd = '', tm = '';
  if (ed && typeof ed.raw === 'number') {
    nd = ed.fmt || ymd_(ed.raw - 4 * 3600);
    const m = Math.round((ed.raw % 86400) / 60); // 協定世界時の分: 13:30 より前は寄付前、19:30 以降は引け後
    tm = m === 0 ? '' : m < 13 * 60 + 30 ? 'BMO' : m >= 19 * 60 + 30 ? 'AMC' : '';
  }
  const rev = num_(ce.revenueAverage);
  const row = [nd, tm, num_(ee.avg != null ? ee.avg : ce.earningsAverage), num_(ee.numberOfAnalysts), num_(ee.yearAgoEps), q0.endDate || '',
    rv(y0, '7daysAgo'), rv(y0, '30daysAgo'), rv(y0, '90daysAgo'), rv(y1, '30daysAgo'), num_(revs.upLast30days), num_(revs.downLast30days),
    hist.map(h => num_(h.surprisePercent)), hist.map(h => String((h.quarter || {}).fmt || '').slice(0, 7)),
    ce.isEarningsDateEstimate ? 1 : 0, rev != null ? rnd_(rev / 1e9, 3) : null];
  if (!nd && row[2] == null && !hist.length && row[7] == null) return null;
  return row;
}

const et_ = ms => Utilities.formatDate(new Date(ms || Date.now()), 'America/New_York', 'yyyy-MM-dd');
const MON_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** 今日から2週間の決算予定（ページの決算カレンダー用。以前の Nasdaq と同じ形） */
function calEarn_(rows) {
  const from = et_(), to = et_(Date.now() + 14 * 86400000), out = [];
  const $ = v => (v == null ? '' : (v < 0 ? '-$' : '$') + Math.abs(v).toFixed(2));
  Object.keys(rows).forEach(t => {
    const r = rows[t];
    if (!r || !r[0] || r[0] < from || r[0] > to) return;
    const fq = r[5] ? MON_[Number(r[5].slice(5, 7)) - 1] + '/' + r[5].slice(0, 4) : '';
    out.push([r[0], t, r[1] === 'BMO' ? 'time-pre-market' : r[1] === 'AMC' ? 'time-after-hours' : 'time-not-supplied', $(r[2]), r[3] != null ? String(r[3]) : '', $(r[4]), fq]);
  });
  return out.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : 1));
}

/** 米国の経済指標の予定と結果（Yahoo Finance の経済指標カレンダー） */
function econEvents_(cr) {
  if (!cr) throw new Error('トークンなし');
  const start = et_(Date.now() - 4 * 86400000), end = et_(Date.now() + 15 * 86400000);
  const ops = [{ operator: 'gte', operands: ['startdatetime', start] }, { operator: 'lte', operands: ['startdatetime', end] }];
  const tries = [ops.concat([{ operator: 'eq', operands: ['country_code', 'US'] }]), ops];
  let last = '';
  for (let k = 0; k < tries.length; k++) {
    const body = { sortType: 'ASC', entityIdType: 'economic_event', sortField: 'startdatetime', offset: 0, size: 250,
      includeFields: ['econ_release', 'country_code', 'startdatetime', 'period', 'after_release_actual', 'consensus_estimate', 'prior_release_actual'],
      query: { operator: 'and', operands: tries[k] } };
    let res;
    try {
      res = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v1/finance/visualization?lang=en-US&region=US&crumb=' + encodeURIComponent(cr.crumb),
        { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), headers: cr.headers, muteHttpExceptions: true });
    } catch (e) { last = msg_(e); continue; }
    if (res.getResponseCode() !== 200) { last = 'http' + res.getResponseCode(); continue; }
    const d0 = ((((JSON.parse(res.getContentText()).finance || {}).result || [])[0] || {}).documents || [])[0] || {};
    const cols = (d0.columns || []).map(c => c.id), ix = id => cols.indexOf(id), out = [];
    const s = v => (v == null ? '' : String(v));
    (d0.rows || []).forEach(row => {
      if (row[ix('country_code')] !== 'US') return;
      const iso = s(row[ix('startdatetime')]), ms = Date.parse(iso);
      if (!ms) return;
      out.push([et_(ms), iso.slice(11, 16), s(row[ix('econ_release')]).replace(/\s*\*\s*$/, '').trim(),
        s(row[ix('after_release_actual')]), s(row[ix('consensus_estimate')]), s(row[ix('prior_release_actual')])]);
    });
    if (out.length) return out;
    last = 'no_rows';
  }
  throw new Error(last || 'no_rows');
}


// ========== ウォッチリストのメール通知 ==========

const KIND = { px: '価格アラート', mv: '値動き', hi: '52週高値', lo: '52週安値', rt: '格付け', tp: '目標株価', er: '決算', rv: '予想修正', test: 'テスト' };
const usd_ = v => (v == null ? '—' : (v < 0 ? '-$' : '$') + Math.abs(v).toFixed(2));
const pct_ = x => (x > 0 ? '+' : '') + (x * 100).toFixed(1) + '%';
const esc_ = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const TICK_ = /^[A-Z0-9][A-Z0-9.\-]{0,11}$/;

function cfg_() { return props_().get('CFG', null); }

/** ページから届いた設定を確かめて保存する */
function saveCfg_(r) {
  const num = v => (typeof v === 'number' && isFinite(v) && v > 0 ? v : null);
  const ty = r.types || {}, watch = (Array.isArray(r.watch) ? r.watch : []).map(t => String(t).toUpperCase()).filter(t => TICK_.test(t)).slice(0, 150);
  const px = {}, names = {};
  Object.keys(r.px || {}).forEach(t => {
    const T = String(t).toUpperCase(), a = num((r.px[t] || {}).a), b = num((r.px[t] || {}).b);
    if (TICK_.test(T) && (a || b)) px[T] = { a: a, b: b };
  });
  watch.forEach(t => { const n = (r.names || {})[t]; if (n) names[t] = String(n).slice(0, 28); });
  const to = String(r.to || '').trim();
  const c = {
    ts: r.ts || Date.now(), on: !!r.on, to: /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to) ? to : '', live: r.live !== false,
    move: Math.min(50, Math.max(1, Number(r.move) || 5)), watch: watch, px: px, names: names, types: {},
  };
  ['px', 'mv', 'hl', 'rt', 'tp', 'er', 'rv'].forEach(k => { c.types[k] = ty[k] !== false; });
  if (JSON.stringify(c).length > 8500) c.names = {};
  props_().set('CFG', c);
  return c;
}

/** 値動き・価格アラート・52週高値/安値を判定する（毎朝の更新と、取引時間中の見張りで共通） */
function evalPrice_(c, P, t, px, prev, hi, lo, day) {
  const items = [], T = c.types, fired = P.get('FIRED', {});
  const a = (c.px[t] || {}).a, b = (c.px[t] || {}).b;
  [['a', a, px >= a, '以上'], ['b', b, px <= b, '以下']].forEach(x => {
    if (!x[1] || !T.px) return;
    const k = x[0] + '|' + t + '|' + x[1];
    if (x[2]) {
      if (!fired[k]) { fired[k] = 1; items.push({ key: 'px|' + k + '|' + Date.now(), t: t, kind: 'px', short: '価格 ' + usd_(px), text: '株価 ' + usd_(px) + ' が、設定した ' + usd_(x[1]) + ' ' + x[3] + 'になりました' }); }
    } else delete fired[k];
  });
  P.set('FIRED', fired);
  if (T.mv && prev) {
    const ch = px / prev - 1;
    if (Math.abs(ch) >= c.move / 100) items.push({ key: 'mv|' + t + '|' + day, t: t, kind: 'mv', short: pct_(ch), text: (ch > 0 ? '上昇 ' : '下落 ') + pct_(ch) + '（' + usd_(px) + '、前日 ' + usd_(prev) + '）' });
  }
  if (T.hl && hi && px > hi) items.push({ key: 'hi|' + t + '|' + day, t: t, kind: 'hi', short: '52週高値', text: '52週高値を更新（' + usd_(px) + '、これまでの高値 ' + usd_(hi) + '）' });
  if (T.hl && lo && px < lo) items.push({ key: 'lo|' + t + '|' + day, t: t, kind: 'lo', short: '52週安値', text: '52週安値を更新（' + usd_(px) + '、これまでの安値 ' + usd_(lo) + '）' });
  return items;
}

/** 毎朝の株価更新のあと: 終値でウォッチリストを判定し、取引時間中の見張りに使う高値・安値・前日終値を残す */
function alertsQuotes_(st, res) {
  const c = cfg_(), P = props_(), ws = {};
  if (!c || !c.watch.length) return;
  let items = [];
  c.watch.forEach(t => {
    const r = res[t];
    if (!r) return;
    const b = r.bars, days = Object.keys(b).sort();
    if (days.length < 3) return;
    const last = days[days.length - 1], past = days.slice(-253, -1);
    const hiP = Math.max.apply(null, past.map(d => b[d][1])), loP = Math.min.apply(null, past.map(d => b[d][2]));
    ws[t] = [Math.max(hiP, b[last][1]), Math.min(loP, b[last][2]), b[last][3], last];
    if (!c.on) return;
    if (r.live) items = items.concat(evalPrice_(c, P, t, r.p, b[last][3], ws[t][0], ws[t][1], r.pt.slice(0, 10)));
    else items = items.concat(evalPrice_(c, P, t, b[last][3], b[days[days.length - 2]][3], hiP, loP, last));
  });
  P.set('WSTATE', ws);
  notify_(st, items);
}

/** アナリスト評価の更新のあと: 格上げ・格下げ、目標株価の変化、決算の前日、EPS予想の修正 */
function alertsAnalyst_(st, prev, rows, earn) {
  const c = cfg_();
  if (!c || !c.on || !c.watch.length) return;
  const T = c.types, today = et_(), since = et_(Date.now() - 4 * 86400000), items = [];
  const d = new Date(Date.parse(today + 'T12:00:00Z'));
  do { d.setUTCDate(d.getUTCDate() + 1); } while (d.getUTCDay() === 0 || d.getUTCDay() === 6);
  const nextBiz = d.toISOString().slice(0, 10), hourET = Number(Utilities.formatDate(new Date(), 'America/New_York', 'H'));
  const ACT = { up: '格上げ', down: '格下げ', init: '新規カバー' }, week = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'YYYY-ww');
  const md = s => Number(s.slice(5, 7)) + '/' + Number(s.slice(8, 10));
  c.watch.forEach(t => {
    const r = rows[t], p = prev[t], e = earn[t];
    if (r && T.rt) (r.ud || []).forEach(u => {
      if (!ACT[u[2]] || u[0] < since) return;
      items.push({ key: 'rt|' + t + '|' + u[0] + '|' + u[1] + '|' + u[2], t: t, kind: 'rt', short: ACT[u[2]],
        text: ACT[u[2]] + '：' + u[1] + '（' + (u[3] ? u[3] + ' → ' : '') + u[4] + '）' + (u[5] ? '・目標 ' + usd_(u[5]) : '') + '・' + md(u[0]) });
    });
    const tn = r && r.t && r.t[0], tp = p && p.t && p.t[0];
    if (T.tp && tn && tp && Math.abs(tn / tp - 1) >= 0.05) {
      items.push({ key: 'tp|' + t + '|' + today, t: t, kind: 'tp', short: '目標株価' + pct_(tn / tp - 1), text: '平均目標株価 ' + usd_(tp) + ' → ' + usd_(tn) + '（' + pct_(tn / tp - 1) + '）' });
    }
    if (e && T.er && e[0] && (e[0] === nextBiz || (e[0] === today && hourET < 16))) {
      const tm = e[1] === 'BMO' ? '寄付前' : e[1] === 'AMC' ? '引け後' : '時間未定';
      items.push({ key: 'er|' + t + '|' + e[0], t: t, kind: 'er', short: '決算 ' + md(e[0]),
        text: '決算発表 ' + md(e[0]) + '（米国時間・' + tm + '）' + (e[2] != null ? '・予想EPS ' + usd_(e[2]) : '') + (e[3] ? '（' + e[3] + '人）' : '') });
    }
    if (e && T.rv && e[6] != null && Math.abs(e[6]) >= 0.03) {
      items.push({ key: 'rv|' + t + '|' + week, t: t, kind: 'rv', short: '予想' + pct_(e[6]),
        text: '今期のEPS予想が7日間で ' + pct_(e[6]) + (e[10] != null ? '（30日の上方修正 ' + e[10] + '件・下方修正 ' + (e[11] || 0) + '件）' : '') });
    }
  });
  notify_(st, items);
}

/** 米国市場の取引時間中（平日 9:30〜16:05 ニューヨーク時間）で、前回の見張りから15分たったか */
function watchDue_(P) {
  const c = cfg_();
  if (!c || !c.on || !c.live || !c.watch.length) return false;
  const u = Utilities.formatDate(new Date(), 'America/New_York', 'u HH:mm').split(' ');
  if (Number(u[0]) > 5 || u[1] < '09:30' || u[1] > '16:05') return false;
  return Date.now() - P.get('WLAST', 0) >= 14.5 * 60000;
}

/** 取引時間中の見張り: ウォッチリストの現在値を取り、値動き・価格アラート・52週高値/安値を判定する */
function watchCheck_(st) {
  const c = cfg_(), P = props_(), ws = P.get('WSTATE', {}), day = et_();
  P.set('WLAST', Date.now());
  const rs = many_(c.watch.map(t => Y + '/v8/finance/chart/' + ysym_(t) + '?range=1d&interval=1d'), HD, Date.now() + 60000, 25);
  let items = [], ok = 0;
  c.watch.forEach((t, i) => {
    let m;
    try { m = json_(rs[i]).chart.result[0].meta; } catch (e) { return; }
    const px = m.regularMarketPrice, prev = m.chartPreviousClose;
    if (!px) return;
    ok++;
    const w = ws[t] && ws[t][3] < day ? ws[t] : null; // 前の取引日までの高値・安値
    items = items.concat(evalPrice_(c, P, t, px, prev, w ? w[0] : null, w ? w[1] : null, day));
  });
  st.live = { at: jst_().at, n: c.watch.length, ok: ok };
  notify_(st, items);
}

/** 条件に合ったものを、まだ知らせていなければ1通にまとめて送る */
function notify_(st, items) {
  const c = cfg_();
  if (!c || !c.on || !items.length) return 0;
  const P = props_(), sent = P.get('SENT', {}), now = jst_(), fresh = [];
  items.forEach(it => { if (!sent[it.key]) { sent[it.key] = now.day; fresh.push(it); } });
  const cut = jst_(Date.now() - 8 * 86400000).day;
  let keys = Object.keys(sent).filter(k => sent[k] >= cut).sort((a, b) => (sent[a] < sent[b] ? -1 : 1));
  while (keys.length && JSON.stringify(keys).length > 7500) keys = keys.slice(1);
  const keep = {};
  keys.forEach(k => { keep[k] = sent[k]; });
  P.set('SENT', keep);
  if (!fresh.length) return 0;
  st.alerts = (st.alerts || []).concat(fresh.map(it => [now.at, it.t, it.kind, it.text])).slice(-80);
  sendMail_(st, c, fresh);
  return fresh.length;
}

function sendMail_(st, c, items, subject) {
  let to = c.to;
  try { to = to || Session.getEffectiveUser().getEmail(); } catch (e) { /* 下で失敗として記録 */ }
  const now = jst_(), names = c.names || {};
  const sub = subject || ('Alpha Lab｜' + items.slice(0, 2).map(it => it.t + ' ' + it.short).join('・') + (items.length > 2 ? ' ほか' + (items.length - 2) + '件' : ''));
  const td = 'padding:10px 8px;border-top:1px solid #E3E6EA;vertical-align:top';
  const html = '<div style="font-family:-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif;color:#111418;max-width:600px">'
    + '<p style="font-size:13px;color:#667080;margin:0 0 10px">Alpha Lab の見張り・' + esc_(now.at) + '（日本時間）</p>'
    + '<table style="border-collapse:collapse;width:100%;font-size:14px;line-height:1.5">'
    + items.map(it => '<tr><td style="' + td + ';white-space:nowrap"><b>' + esc_(it.t) + '</b>' + (names[it.t] ? '<br><span style="color:#667080;font-size:12px">' + esc_(names[it.t]) + '</span>' : '') + '</td>'
      + '<td style="' + td + '"><span style="font-size:11px;color:#667080">' + esc_(KIND[it.kind] || '') + '</span><br>' + esc_(it.text) + '</td></tr>').join('')
    + '</table><p style="margin:16px 0 0"><a href="' + PAGE + '" style="color:#2B48D6">Alpha Lab を開く</a></p>'
    + '<p style="font-size:11px;color:#9098A3;margin-top:14px">Google Apps Script から自動で送っています。止めるには Alpha Lab の「メール通知」をオフにしてください。</p></div>';
  try {
    if (!to) throw new Error('送信先がありません');
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('今日の送信上限に達しました');
    MailApp.sendEmail({ to: to, subject: sub, htmlBody: html, name: 'Alpha Lab' });
    st.mail = { at: now.at, ok: true, to: to, n: items.length };
  } catch (e) {
    st.mail = { at: now.at, ok: false, to: to || '', err: msg_(e) };
  }
}

/** ページの「テスト送信」 */
function testMail_(st) {
  const c = cfg_() || { watch: [], names: {}, types: {} };
  sendMail_(st, c, [{ t: 'Alpha Lab', kind: 'test', short: '', text: 'メール通知のテストです。ウォッチリスト ' + (c.watch || []).length + '銘柄を見張っています'
    + (c.on ? '。' : '（通知は今オフです）。') }], 'Alpha Lab｜メール通知のテスト');
}


// ========== 接続テスト ==========

function diagnose_() {
  const probe = (name, url, hd, ok) => {
    try {
      const r = UrlFetchApp.fetch(url, { headers: hd || HD, muteHttpExceptions: true });
      const code = r.getResponseCode();
      return [name, code === 200 && (!ok || ok(r.getContentText())) ? 'ok' : 'http' + code];
    } catch (e) { return [name, msg_(e)]; }
  };
  const out = {};
  [
    probe('Yahoo 株価', Y + '/v8/finance/chart/AAPL?range=5d&interval=1d', HD, t => t.indexOf('"result"') >= 0),
    probe('Yahoo ニュース', Y + '/v1/finance/search?q=AAPL&quotesCount=0&newsCount=2', HD),
    probe('FRED 経済データ', fredUrl_('DFF', 1), HD, t => t.indexOf('DFF') >= 0),
    probe('CNN Fear & Greed', 'https://production.dataviz.cnn.io/index/fearandgreed/graphdata/' + jst_().day, CNN),
  ].forEach(p => { out[p[0]] = p[1]; });
  out['Yahoo トークン'] = crumb_() ? 'ok' : '取得できず（アナリストのコンセンサスは調査会社の評価で代替）';
  return out;
}
