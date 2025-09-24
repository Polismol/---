/*********************************************
 * LiveDune → Контент-план ТОС (лист "Контент-план")
 * Линки: "Ссылка ТГ" / "Ссылка ВК" / "Ссылка ОК"
 * Метрики: TG → "просмотры","интеракции","ER"
 *          VK → "охват","интеракции.1","ER.1"
 *          OK → "охват.1","интеракции.2","ER.2"
 *********************************************/

const SHEET_NAME = 'Контент-план';

const COLS = {
  date:        'Дата',

  tg_url:      'Ссылка ТГ',
  tg_views:    'просмотры',
  tg_inter:    'интеракции',
  tg_er:       'ER',

  vk_url:      'Ссылка ВК',
  vk_reach:    'охват',
  vk_inter:    'интеракции.1',
  vk_er:       'ER.1',

  ok_url:      'Ссылка ОК',
  ok_views:    'охват.1',       // используем как «просмотры» для OK
  ok_inter:    'интеракции.2',
  ok_er:       'ER.2'
};

// ==== Меню ====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Статистика')
    .addSubMenu(
      ui.createMenu('Собрать показатели')
        .addItem('Январь',   'collectMonth_1')
        .addItem('Февраль',  'collectMonth_2')
        .addItem('Март',     'collectMonth_3')
        .addItem('Апрель',   'collectMonth_4')
        .addItem('Май',      'collectMonth_5')
        .addItem('Июнь',     'collectMonth_6')
        .addItem('Июль',     'collectMonth_7')
        .addItem('Август',   'collectMonth_8')
        .addItem('Сентябрь', 'collectMonth_9')
        .addItem('Октябрь',  'collectMonth_10')
        .addItem('Ноябрь',   'collectMonth_11')
        .addItem('Декабрь',  'collectMonth_12')
    )
    .addSeparator()
    .addItem('Собрать за ВСЮ таблицу', 'collectAll')
    .addToUi();
}

// ==== Команды ====
function collectAll() {
  const sheet = getSheet_();
  const meta = buildHeaderIndex_(sheet);
  const token = getToken_();
  const accIndex = getAccountsIndex_(token);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();

  for (let i = 0; i < values.length; i++) {
    processRow_(sheet, 2 + i, values[i], meta, token, accIndex);
  }
  SpreadsheetApp.getActive().toast('Готово: обновили все строки с ссылками.');
}

function collectMonth_(month) {
  const sheet = getSheet_();
  const meta = buildHeaderIndex_(sheet);
  const token = getToken_();
  const accIndex = getAccountsIndex_(token);

  const year = new Date().getFullYear();
  const from = new Date(year, month - 1, 1);
  const to   = new Date(year, month, 1);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const d = parseDateSafe_(row[meta.colIdx.date]);
    if (d && d >= from && d < to) {
      processRow_(sheet, 2 + i, row, meta, token, accIndex);
    }
  }
  SpreadsheetApp.getActive().toast(`Готово: обновили за ${from.toLocaleDateString()}–${new Date(to - 1).toLocaleDateString()}.`);
}
function collectMonth_1(){ collectMonth_(1); }
function collectMonth_2(){ collectMonth_(2); }
function collectMonth_3(){ collectMonth_(3); }
function collectMonth_4(){ collectMonth_(4); }
function collectMonth_5(){ collectMonth_(5); }
function collectMonth_6(){ collectMonth_(6); }
function collectMonth_7(){ collectMonth_(7); }
function collectMonth_8(){ collectMonth_(8); }
function collectMonth_9(){ collectMonth_(9); }
function collectMonth_10(){ collectMonth_(10); }
function collectMonth_11(){ collectMonth_(11); }
function collectMonth_12(){ collectMonth_(12); }

// ==== Обработка строки ====
function processRow_(sheet, rowNumber, rowValues, meta, token, accIndex) {
  // Telegram
  const tgUrl = getCellStr_(rowValues, meta.colIdx.tg_url);
  if (tgUrl) {
    try {
      const m = fetchMetricsForUrl_(cleanUrl_(tgUrl), 'telegram', token, accIndex);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.tg_views + 1, m.views);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.tg_inter + 1, m.interactions);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.tg_er    + 1, m.er);
    } catch(e) {
      writeIfChanged_(sheet, rowNumber, meta.colIdx.tg_views + 1, 'ERR: ' + (e.message || e));
    }
  }

  // VK
  const vkUrl = getCellStr_(rowValues, meta.colIdx.vk_url);
  if (vkUrl) {
    try {
      const m = fetchMetricsForUrl_(cleanUrl_(vkUrl), 'vk', token, accIndex);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.vk_reach + 1, m.reach);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.vk_inter + 1, m.interactions);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.vk_er    + 1, m.er);
    } catch(e) {
      writeIfChanged_(sheet, rowNumber, meta.colIdx.vk_reach + 1, 'ERR: ' + (e.message || e));
    }
  }

  // OK
  const okUrl = getCellStr_(rowValues, meta.colIdx.ok_url);
  if (okUrl) {
    try {
      const m = fetchMetricsForUrl_(cleanUrl_(okUrl), 'ok', token, accIndex);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.ok_views + 1, m.views);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.ok_inter + 1, m.interactions);
      writeIfChanged_(sheet, rowNumber, meta.colIdx.ok_er    + 1, m.er);
    } catch(e) {
      writeIfChanged_(sheet, rowNumber, meta.colIdx.ok_views + 1, 'ERR: ' + (e.message || e));
    }
  }
}

// ==== LiveDune: метрики по ссылке ====
function fetchMetricsForUrl_(url, platform, token, accIndex) {
  const parsed = parseUrl_(url, platform);
  if (!parsed) throw new Error('Не удалось разобрать ссылку: ' + url);

  const accountId = resolveAccountId_(platform, parsed, accIndex);
  if (!accountId) throw new Error(`Аккаунт не найден в LiveDune для: ${url}`);

  const post = getPostStat_(accountId, parsed.postId, token);
  if (!post) throw new Error('Пост не найден в LiveDune: ' + url);

  const reactions = post.reactions || {};
  let interactions = 0;
  Object.keys(reactions).forEach(k => {
    const v = reactions[k];
    if (typeof v === 'number') interactions += v;
    else if (typeof v === 'string' && !isNaN(+v)) interactions += +v;
  });

  let base = null;
  if (platform === 'vk') {
    base = tryNum_(post?.reach?.total) ?? tryNum_(post?.impressions?.total);
  } else if (platform === 'telegram') {
    base = tryNum_(post?.impressions?.total) ?? tryNum_(post?.views) ?? tryNum_(post?.video_views);
  } else if (platform === 'ok') {
    base = tryNum_(post?.impressions?.total) ?? tryNum_(post?.views);
  }

  const out = { interactions: interactions || 0, er: 0 };
  if (platform === 'vk') out.reach = base || 0; else out.views = base || 0;
  out.er = base && base > 0 ? +((interactions / base) * 100).toFixed(2) : 0;
  return out;
}

// ==== Парсинг URL (регулярки прежде всего) ====
function parseUrl_(rawUrl, platformHint) {
  if (!rawUrl) return null;

  // Нормализуем «грязные» символы, дефисы и пробелы
  let s = (rawUrl || '').toString().trim()
    .replace(/[\u200B-\u200D\uFEFF]/g, '')   // zero-width
    .replace(/\u00A0/g, ' ')                 // NBSP → space
    .replace(/[\u2010-\u2015\u2212]/g, '-'); // экзотические дефисы → '-'

  // Telegram
  if (platformHint === 'telegram' || /(?:^|\.)t(?:elegram)?\.me/i.test(s)) {
    let m = s.match(/(?:https?:\/\/)?(?:t\.me|telegram\.me)\/(?:s\/)?([A-Za-z0-9_]+)\/(\d+)/i);
    if (m) return { platform: 'telegram', postId: m[2], baseUrl: `https://t.me/${m[1]}` };
    m = s.match(/(?:https?:\/\/)?(?:t\.me|telegram\.me)\/c\/(\d+)\/(\d+)/i);
    if (m) return { platform: 'telegram', postId: m[2], baseUrl: `https://t.me/c/${m[1]}` };
    return null;
  }

  // VK
  if (platformHint === 'vk' || /(?:^|\.)vk\.com/i.test(s)) {
    let m = s.match(/vk\.com\/wall-?(\d+)_(\d+)/i);
    if (m) {
      const socialGroupId = m[1], postId = m[2];
      const name = (s.match(/vk\.com\/([A-Za-z0-9_\.]+)(?:[?#]|$)/i) || [])[1];
      const baseUrl = name && !/^wall-?\d+_\d+$/i.test(name) && !/^club\d+$/i.test(name) && !/^public\d+$/i.test(name)
        ? `https://vk.com/${name}` : `https://vk.com/club${socialGroupId}`;
      return { platform: 'vk', postId, socialGroupId, baseUrl };
    }
    m = s.match(/[?&#](?:w|z)=wall-?(\d+)_(\d+)/i);
    if (m) {
      const socialGroupId = m[1], postId = m[2];
      const name = (s.match(/vk\.com\/([A-Za-z0-9_\.]+)(?:[?#]|$)/i) || [])[1];
      const baseUrl = name && !/^wall-?\d+_\d+$/i.test(name) && !/^club\d+$/i.test(name) && !/^public\d+$/i.test(name)
        ? `https://vk.com/${name}` : `https://vk.com/club${socialGroupId}`;
      return { platform: 'vk', postId, socialGroupId, baseUrl };
    }
    return null;
  }

  // OK
  if (platformHint === 'ok' || /(?:^|\.)ok\.ru/i.test(s)) {
    const m = s.match(/ok\.ru\/(?:[^/?#]+\/)?topic\/(\d+)/i);
    if (!m) return null;
    const postId = m[1];
    const base = (s.match(/(ok\.ru\/[^?#]+?)\/topic\/\d+/i) || [])[1];
    const baseUrl = base ? `https://` + base.replace(/^https?:\/\//i, '').replace(/^www\./i,'') : 'https://ok.ru';
    return { platform: 'ok', postId, baseUrl };
  }

  // Fallback через URL API
  try {
    if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
    const u = new URL(s);
    const host = u.hostname.replace(/^www\./i,'').replace(/^m\./i,'').toLowerCase();
    const segs = (u.pathname || '/').replace(/^\/+/, '').split('/').filter(Boolean);

    if ((platformHint === 'telegram' || /(t|telegram)\.me$/i.test(host)) && segs.length >= 2) {
      if (segs[0].toLowerCase() === 's' && segs.length >= 3)
        return { platform: 'telegram', postId: segs[2], baseUrl: `https://t.me/${segs[1]}` };
      if (segs[0].toLowerCase() === 'c' && segs.length >= 3)
        return { platform: 'telegram', postId: segs[2], baseUrl: `https://t.me/c/${segs[1]}` };
      return { platform: 'telegram', postId: segs[1], baseUrl: `https://t.me/${segs[0]}` };
    }
    if ((platformHint === 'vk' || /vk\.com$/i.test(host))) {
      const mq = (u.searchParams.get('w') || u.searchParams.get('z') || '').match(/wall-?(\d+)_(\d+)/i);
      const mp = u.pathname.match(/\/wall-?(\d+)_(\d+)/i);
      const m  = mp || mq;
      if (!m) return null;
      const socialGroupId = m[1], postId = m[2];
      const first = segs[0] || '';
      const baseUrl = first && !/^wall-?\d+_\d+$/i.test(first) && !/^club\d+$/i.test(first) && !/^public\d+$/i.test(first)
        ? `https://vk.com/${first}` : `https://vk.com/club${socialGroupId}`;
      return { platform: 'vk', postId, socialGroupId, baseUrl };
    }
    if ((platformHint === 'ok' || /ok\.ru$/i.test(host))) {
      const mt = u.pathname.match(/\/topic\/(\d+)/i);
      if (!mt) return null;
      const postId = mt[1];
      const idx = segs.findIndex(sg => sg.toLowerCase() === 'topic');
      const baseUrl = idx > 0 ? `https://ok.ru/${segs.slice(0, idx).join('/')}` : 'https://ok.ru';
      return { platform: 'ok', postId, baseUrl };
    }
  } catch (_) {}

  return null;
}

// ==== Соответствие аккаунту LiveDune ====
function resolveAccountId_(platform, parsed, accIndex) {
  const keyType = platform === 'vk' ? 'vk_group'
                 : platform === 'ok' ? 'ok_group'
                 : platform === 'telegram' ? 'telegram'
                 : null;
  if (!keyType) return null;

  const list = accIndex[keyType] || [];

  const direct = list.find(a => (a.url || '').toLowerCase().startsWith(parsed.baseUrl.toLowerCase()));
  if (direct) return direct.id;

  if (platform === 'vk' && parsed.socialGroupId) {
    const byId = list.find(a => String(a.social_id) === String(parsed.socialGroupId));
    if (byId) return byId.id;
  }

  const host = parsed.baseUrl.split('/')[2];
  const byHost = list.find(a => (a.url || '').includes(host));
  return byHost ? byHost.id : null;
}

// ==== LiveDune API ====
function getPostStat_(accountId, postId, token) {
  const url = `https://api.livedune.com/accounts/${encodeURIComponent(accountId)}/posts/${encodeURIComponent(postId)}?access_token=${encodeURIComponent(token)}`;
  const res = httpGetJson_(url);
  const arr = res && res.response;
  return Array.isArray(arr) && arr.length ? arr[0] : null;
}

function getAccountsIndex_(token) {
  const out = { vk_group: [], ok_group: [], telegram: [] };
  let after = null, guard = 0;
  do {
    const url = `https://api.livedune.com/accounts?access_token=${encodeURIComponent(token)}${after ? `&after=${after}` : ''}`;
    const res = httpGetJson_(url);
    const list = (res && res.response) || [];
    list.forEach(a => { if (out[a.type]) out[a.type].push(a); });
    after = res && res.after;
    guard++;
  } while (after && guard < 20);
  return out;
}

function httpGetJson_(url) {
  const opts = { method: 'get', muteHttpExceptions: true, followRedirects: true };
  let attempt = 0, wait = 600;
  while (attempt < 5) {
    const resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      const txt = resp.getContentText();
      try { return JSON.parse(txt); } catch { throw new Error('Bad JSON from LiveDune'); }
    }
    Utilities.sleep(wait);
    wait = Math.min(wait * 2, 8000);
    attempt++;
  }
  throw new Error('LiveDune API error: ' + url);
}

// ==== Утилиты: шапка, даты, запись, нормализация ====
function getSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Лист "${SHEET_NAME}" не найден`);
  return sh;
}

function norm_(s) {
  return (s || '')
    .toString()
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/[.·•●]/g, '.')
    .trim()
    .toLowerCase();
}

function findHeaderIdx_(headerArr, preferred, aliases, fallbackRegex) {
  const nHeader = headerArr.map(norm_);
  const targets = [preferred].concat(aliases || []).map(norm_);
  for (const t of targets) {
    const i = nHeader.indexOf(t);
    if (i !== -1) return i;
  }
  if (fallbackRegex) {
    for (let i = 0; i < nHeader.length; i++) {
      if (fallbackRegex.test(nHeader[i])) return i;
    }
  }
  return -1;
}

function buildHeaderIndex_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const header = headerRange.getValues()[0].map(v => (v || '').toString());

  const idx_date     = findHeaderIdx_(header, COLS.date, [], /^дата$/);

  const idx_tg_url   = findHeaderIdx_(header, COLS.tg_url, ['ссылка тг','tg','телеграм'], /ссыл.*тг|t(\.|e)?le?gra?m/i);
  const idx_tg_views = findHeaderIdx_(header, COLS.tg_views, ['просмотры тг'], /просмотры(\s*tg)?$/i);
  const idx_tg_inter = findHeaderIdx_(header, COLS.tg_inter, ['инетракции','интеракции тг'], /ин(е|и)т?ракц/i);
  const idx_tg_er    = findHeaderIdx_(header, COLS.tg_er, ['er тг','er tg'], /^er(\s*tg)?$/i);

  const idx_vk_url   = findHeaderIdx_(header, COLS.vk_url, ['ссылка вк','vk'], /ссыл.*вк|vk/i);
  const idx_vk_reach = findHeaderIdx_(header, COLS.vk_reach, ['охват вк','reach'], /охват|reach/i);
  const idx_vk_inter = findHeaderIdx_(header, COLS.vk_inter, ['инетракции.1','интеракции вк','интеракции 1','интеракции .1'], /ин(е|и)т?ракц.*(\.| )?1$/i);
  const idx_vk_er    = findHeaderIdx_(header, COLS.vk_er, ['er.1','er вк'], /^er(\.| )?1$/i);

  const idx_ok_url   = findHeaderIdx_(header, COLS.ok_url, ['ссылка ок','ok','одноклассники'], /ссыл.*ok|однокл/i);
  const idx_ok_views = findHeaderIdx_(header, COLS.ok_views, ['просмотры ок','impressions'], /(просмотры|impressions|охват\.?1)$/i);
  const idx_ok_inter = findHeaderIdx_(header, COLS.ok_inter, ['интеракции.2','инетракции.2','интеракции ок','интеракции 2','интеракции .2'], /ин(е|и)т?ракц.*(\.| )?2$/i);
  const idx_ok_er    = findHeaderIdx_(header, COLS.ok_er, ['er.2','er ок'], /^er(\.| )?2$/i);

  const miss = [];
  function req(name, idx) { if (idx === -1) miss.push(name); }
  req(COLS.date, idx_date);
  req(COLS.tg_url, idx_tg_url);   req(COLS.tg_views, idx_tg_views); req(COLS.tg_inter, idx_tg_inter); req(COLS.tg_er, idx_tg_er);
  req(COLS.vk_url, idx_vk_url);   req(COLS.vk_reach, idx_vk_reach); req(COLS.vk_inter, idx_vk_inter); req(COLS.vk_er, idx_vk_er);
  req(COLS.ok_url, idx_ok_url);   req(COLS.ok_views, idx_ok_views); req(COLS.ok_inter, idx_ok_inter); req(COLS.ok_er, idx_ok_er);

  if (miss.length) {
    const seen = header.map(h => `• ${h}`).join('\n');
    throw new Error('Не нашёл колонки:\n' + miss.join(', ') + '\n\nЧто вижу в шапке:\n' + seen);
  }

  return {
    header,
    colIdx: {
      date:     idx_date,
      tg_url:   idx_tg_url,  tg_views: idx_tg_views, tg_inter: idx_tg_inter, tg_er: idx_tg_er,
      vk_url:   idx_vk_url,  vk_reach: idx_vk_reach, vk_inter: idx_vk_inter, vk_er: idx_vk_er,
      ok_url:   idx_ok_url,  ok_views: idx_ok_views, ok_inter: idx_ok_inter, ok_er: idx_ok_er,
    }
  };
}

function writeIfChanged_(sheet, row, col, val) {
  const rng = sheet.getRange(row, col);
  const cur = rng.getValue();
  if (String(cur) !== String(val)) rng.setValue(val);
}

function getCellStr_(rowValues, idx) {
  if (idx == null || idx < 0) return '';
  const v = rowValues[idx];
  return (v == null) ? '' : String(v).trim();
}

function parseDateSafe_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const n = Date.parse(v);
  return isNaN(n) ? null : new Date(n);
}

function tryNum_(x) {
  if (x == null) return null;
  const n = +x;
  return isNaN(n) ? null : n;
}

// НОРМАЛИЗАЦИЯ URL: убираем zero-width/NBSP, приводим дефисы, унифицируем хост
function cleanUrl_(u) {
  let s = (u || '').toString().trim()
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u2010-\u2015\u2212]/g, '-');
  if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
  try {
    const url = new URL(s);
    url.hostname = url.hostname
      .replace(/^www\./i, '')
      .replace(/^m\./i, '')
      .replace(/^telegram\.me$/i, 't.me');
    return url.toString();
  } catch {
    return s;
  }
}

// ==== Токен LiveDune: выполнить один раз ====
function saveToken() {
  const token = 'aa68777b612ce654.54205968';
  PropertiesService.getScriptProperties().setProperty('LIVEDUNE_TOKEN', token);
  SpreadsheetApp.getActive().toast('Токен LiveDune сохранён.');
}
function getToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('LIVEDUNE_TOKEN');
  if (!token) throw new Error('Не задан токен LiveDune. Запусти saveToken() один раз.');
  return token;
}
