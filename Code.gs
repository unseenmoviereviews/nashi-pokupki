/**
 * Наши покупочки — серверная часть (Google Apps Script).
 * Каталог того, что мы уже купили, с оценками по критериям.
 * Сайт лежит на GitHub Pages и обращается сюда как к API. Вход — по личному ключу,
 * ключи хранятся только в свойствах скрипта (в коде и на GitHub их нет).
 *
 * 1. Впишите имена в MEMBERS и (по желанию) почту пары в SHARE_WITH.
 * 2. Один раз запустите функцию setup() — в журнале появятся ключи входа.
 * 3. Разверните как веб-приложение (см. README.md).
 */
// Стартовые имена. Потом каждый меняет своё имя в приложении — актуальный список хранится
// в свойствах скрипта, и правка этой строки на уже работающее приложение не влияет.
const MEMBERS = ['Я', 'Сонюшка'];

// ID таблицы (из её адреса: /spreadsheets/d/ЭТОТ_ID/edit). Нужен, только если скрипт
// создан отдельно, а не через «Расширения → Apps Script» в самой таблице.
const SPREADSHEET_ID = '';

// Кому открыть саму таблицу (чтобы смотреть её напрямую). Можно оставить пустым.
const SHARE_WITH = [];

// Варианты ответа на «Купили бы снова?»
const AGAIN = ['Да', 'Возможно', 'Нет'];

// Стартовые критерии. Потом их можно менять прямо на листе «Критерии».
const DEFAULT_CRITERIA = [
  ['Одежда', 'Качество, Посадка, Внешний вид, Удобство, Цена/качество'],
  ['Обувь', 'Удобство, Качество, Внешний вид, Цена/качество'],
  ['Техника', 'Функциональность, Качество сборки, Дизайн, Удобство, Цена/качество'],
  ['Дом', 'Качество, Внешний вид, Практичность, Цена/качество'],
  ['Косметика', 'Эффект, Запах, Текстура, Цена/качество'],
  ['Еда', 'Вкус, Качество, Цена/качество'],
  ['Другое', 'Качество, Внешний вид, Полезность, Цена/качество'],
];

const SHEETS = { items: 'Товары', reviews: 'Оценки', comments: 'Комментарии', criteria: 'Критерии' };

// [ключ в коде, заголовок столбца]. Столбцы с "_" — служебные, они скрыты.
const SCHEMA = {
  items: [['id', 'ID'], ['created', 'Добавлено'], ['author', 'Автор'], ['name', 'Название'],
    ['url', 'Ссылка'], ['shop', 'Магазин'], ['bought', 'Дата покупки'], ['price', 'Цена'], ['currency', 'Валюта'],
    ['category', 'Категория'], ['again', 'Купили бы снова'], ['description', 'Описание'],
    ['avg', 'Средняя оценка'], ['photoLinks', 'Фото'], ['photos', '_photos'],
    ['offersText', 'Цены в других магазинах'], ['offers', '_offers']],
  reviews: [['itemId', 'ID товара'], ['itemName', 'Товар'], ['author', 'Автор'], ['updated', 'Обновлено'],
    ['avg', 'Средняя'], ['ratingsText', 'Оценки'], ['ratings', '_ratings'], ['again', 'Купил(а) бы снова']],
  comments: [['id', 'ID'], ['itemId', 'ID товара'], ['itemName', 'Товар'], ['author', 'Автор'],
    ['date', 'Дата'], ['text', 'Текст']],
  criteria: [['category', 'Категория'], ['list', 'Критерии (через запятую)']],
};

/* ---------- API ---------- */

// Функции, которые может вызывать сайт.
const API = { getData, saveItem, deleteItem, saveReview, addComment, deleteComment, uploadPhoto, getThumbs, getPhoto, renameMe,
  savePush, removePush, getNews, testPush };
let currentUser = null;

function doGet() {
  return ContentService.createTextOutput('API «Наши покупочки» работает.');
}

// Запрос: {key, fn, args}. Ответ: {ok: true, data} или {ok: false, error}.
function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    currentUser = userByKey_(req.key);
    if (!currentUser) throw new Error('Неверный ключ');
    if (!Object.prototype.hasOwnProperty.call(API, req.fn)) throw new Error('Неизвестная функция');
    res = { ok: true, data: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Первичная настройка (запустить один раз) ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || (SPREADSHEET_ID && SpreadsheetApp.openById(SPREADSHEET_ID));
  if (!ss) throw new Error('Не нашёл таблицу: впишите SPREADSHEET_ID в начале файла.');
  Object.keys(SHEETS).forEach(key => {
    const sh = ss.getSheetByName(SHEETS[key]) || ss.insertSheet(SHEETS[key]);
    if (sh.getLastRow() === 0) {
      const headers = SCHEMA[key].map(c => c[1]);
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      headers.forEach((h, i) => { if (h[0] === '_') sh.hideColumns(i + 1); });
    }
  });
  const crit = ss.getSheetByName(SHEETS.criteria);
  if (crit.getLastRow() === 1) crit.getRange(2, 1, DEFAULT_CRITERIA.length, 2).setValues(DEFAULT_CRITERIA);
  const blank = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  let folderId = props.getProperty('FOLDER_ID');
  if (!folderId) {
    folderId = DriveApp.createFolder('Наши покупки — фото').getId();
    props.setProperty('FOLDER_ID', folderId);
  }
  SHARE_WITH.forEach(email => ss.addViewer(email));

  // Личные ключи входа: создаются один раз, повторный запуск их не меняет.
  members_().forEach(name => {
    if (!props.getProperty('KEY_' + name)) props.setProperty('KEY_' + name, Utilities.getUuid().replace(/-/g, ''));
  });
  showKeys();
}

/** Печатает ключи входа в журнал выполнения. */
function showKeys() {
  const props = PropertiesService.getScriptProperties();
  members_().forEach(name => Logger.log(name + ': ' + props.getProperty('KEY_' + name)));
}

/** Если ключ утёк: запустите, затем showKeys() и войдите заново с новыми ключами. */
function resetKeys() {
  const props = PropertiesService.getScriptProperties();
  members_().forEach(name => props.setProperty('KEY_' + name, Utilities.getUuid().replace(/-/g, '')));
  showKeys();
}

/* ---------- API для страницы ---------- */

function getData() {
  const user = me_();
  const criteria = {};
  read_('criteria').forEach(c => {
    if (c.category) criteria[c.category] = String(c.list).split(',').map(s => s.trim()).filter(String);
  });
  const byId = {};
  const items = read_('items').map(it => {
    const out = {
      id: it.id, created: it.created, author: it.author, name: it.name, url: it.url, shop: it.shop,
      bought: it.bought ? String(it.bought).slice(0, 10) : '',
      price: it.price, currency: it.currency, category: it.category, again: it.again,
      description: it.description, photos: parse_(it.photos, []), offers: parse_(it.offers, []), reviews: [], comments: [],
    };
    byId[it.id] = out;
    return out;
  });
  read_('reviews').forEach(r => {
    if (byId[r.itemId]) byId[r.itemId].reviews.push({ author: r.author, updated: r.updated, ratings: parse_(r.ratings, {}), again: r.again || '' });
  });
  read_('comments').forEach(c => {
    if (byId[c.itemId]) byId[c.itemId].comments.push({ id: c.id, author: c.author, date: c.date, text: c.text });
  });
  return { me: user.name, members: members_(), criteria, again: AGAIN, items, pushKey: vapidKeys_().pub };
}

function saveItem(d) {
  const user = me_();
  const id = withLock_(() => {
    let it = null;
    if (d.id) {
      it = read_('items').find(x => x.id === d.id);
      if (!it) throw new Error('Товар не найден — возможно, его удалили.');
      it.created = new Date(it.created);
    } else {
      it = { id: newId_(), created: new Date(), author: user.name, avg: '', again: '' };
    }
    // «Купили бы снова» у каждого свой и сохраняется вместе с оценкой (saveReview).
    ['name', 'url', 'shop', 'currency', 'category', 'description'].forEach(k => {
      it[k] = String(d[k] || '').trim();
    });
    // Дата приходит как ГГГГ-ММ-ДД; полдень — чтобы часовой пояс не сдвинул день.
    it.bought = /^\d{4}-\d{2}-\d{2}$/.test(d.bought || '') ? new Date(d.bought + 'T12:00:00') : '';
    if (!it.name) throw new Error('Укажите название');
    it.price = d.price === '' || d.price == null || isNaN(Number(d.price)) ? '' : Number(d.price);

    const photos = (d.photos || []).map(p => ({ full: String(p.full), thumb: String(p.thumb) }));
    parse_(it.photos, []).filter(p => !photos.some(q => q.full === p.full)).forEach(trashPhoto_);
    it.photos = JSON.stringify(photos);
    it.photoLinks = photos.map(p => 'https://drive.google.com/file/d/' + p.full + '/view').join('\n');

    // Цены в других магазинах — чтобы видеть, где дешевле.
    const offers = (d.offers || []).map(o => ({
      shop: String(o.shop || '').trim().slice(0, 60),
      price: o.price === '' || o.price == null || isNaN(Number(o.price)) ? '' : Number(o.price),
      currency: String(o.currency || '').trim().slice(0, 4),
      url: String(o.url || '').trim().slice(0, 500),
    })).filter(o => o.shop || o.price !== '').slice(0, 10);
    it.offers = JSON.stringify(offers);
    it.offersText = offers.map(o => (o.shop || '?') + ': ' + (o.price !== '' ? o.price + ' ' + o.currency : '—')).join('; ');
    write_('items', it);
    return it.id;
  });
  if (!d.id) notify_({ type: 'item', author: user.name, itemId: id, name: String(d.name || '').trim() });
  return id;
}

function deleteItem(id) {
  me_();
  return withLock_(() => {
    const it = read_('items').find(x => x.id === id);
    if (!it) return;
    parse_(it.photos, []).forEach(trashPhoto_);
    deleteRows_('reviews', r => r.itemId === id);
    deleteRows_('comments', c => c.itemId === id);
    sheet_('items').deleteRow(it._row);
  });
}

/** Сохраняет оценку текущего пользователя и его «купил(а) бы снова» (по одной записи на человека на товар). */
function saveReview(itemId, ratings, again) {
  const user = me_();
  const news = withLock_(() => {
    const it = read_('items').find(x => x.id === itemId);
    if (!it) throw new Error('Товар не найден');
    const clean = {};
    Object.keys(ratings || {}).forEach(k => {
      const v = Math.round(Number(ratings[k]));
      if (v >= 1 && v <= 5) clean[String(k)] = v;
    });
    again = AGAIN.indexOf(again) >= 0 ? again : '';
    const reviews = read_('reviews').filter(r => r.itemId === itemId);
    const mine = reviews.find(r => r.author === user.name);
    if (Object.keys(clean).length || again) {
      write_('reviews', {
        _row: mine && mine._row, itemId, itemName: it.name, author: user.name, updated: new Date(),
        avg: avg_(Object.values(clean)), ratingsText: Object.keys(clean).map(k => k + ': ' + clean[k]).join(', '),
        ratings: JSON.stringify(clean), again,
      });
    } else if (mine) {
      sheet_('reviews').deleteRow(mine._row);
    }
    // Пересчитываем общую среднюю товара и сводку «купили бы снова» по всем участникам.
    const others = reviews.filter(r => r.author !== user.name);
    const all = [];
    others.forEach(r => all.push(...Object.values(parse_(r.ratings, {}))));
    all.push(...Object.values(clean));
    it.avg = avg_(all);
    const answers = others.filter(r => r.again).map(r => r.author + ': ' + r.again);
    if (again) answers.push(user.name + ': ' + again);
    it.again = answers.join(', ');
    it.created = new Date(it.created);
    if (it.bought) it.bought = new Date(it.bought);
    write_('items', it);
    // Уведомляем только о первой оценке чужой покупки: о своей партнёр уже узнал, когда она появилась.
    if (mine || it.author === user.name || !(Object.keys(clean).length || again)) return null;
    const score = avg_(Object.values(clean));
    return { type: 'review', author: user.name, itemId, name: it.name,
      text: [score !== '' ? score + ' из 5' : '', again ? 'снова: ' + again.toLowerCase() : ''].filter(String).join(', ') };
  });
  if (news) notify_(news);
}

function addComment(itemId, text) {
  const user = me_();
  text = String(text || '').trim();
  if (!text) return;
  const name = withLock_(() => {
    const it = read_('items').find(x => x.id === itemId);
    if (!it) throw new Error('Товар не найден');
    write_('comments', { id: newId_(), itemId, itemName: it.name, author: user.name, date: new Date(), text });
    return it.name;
  });
  notify_({ type: 'comment', author: user.name, itemId, name, text: text.slice(0, 120) });
}

function deleteComment(id) {
  const user = me_();
  return withLock_(() => {
    deleteRows_('comments', c => c.id === id && c.author === user.name);
  });
}

/** Принимает фото (base64 JPEG, уже уменьшенное на телефоне) и миниатюру. */
function uploadPhoto(fullB64, thumbB64) {
  me_();
  const folder = DriveApp.getFolderById(prop_('FOLDER_ID'));
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss') + '_' + newId_();
  const full = folder.createFile(Utilities.newBlob(Utilities.base64Decode(fullB64), 'image/jpeg', stamp + '.jpg'));
  const thumb = folder.createFile(Utilities.newBlob(Utilities.base64Decode(thumbB64), 'image/jpeg', stamp + '_mini.jpg'));
  return { full: full.getId(), thumb: thumb.getId() };
}

/** Отдаёт миниатюры как data-URL: файлы приватные, прямые ссылки на них страница показать не может. */
function getThumbs(ids) {
  me_();
  const out = {};
  (ids || []).slice(0, 30).forEach(id => {
    try {
      out[id] = 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes());
    } catch (e) {
      out[id] = null;
    }
  });
  return out;
}

/** Фото в полном размере для просмотра на сайте. */
function getPhoto(id) {
  me_();
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes());
}

/**
 * Меняет имя текущего пользователя: в списке участников, у ключа входа
 * и в столбце «Автор» на всех листах, чтобы старые записи остались его.
 */
function renameMe(newName) {
  const user = me_();
  newName = String(newName || '').replace(/\s+/g, ' ').trim();
  if (!newName) throw new Error('Имя не может быть пустым');
  if (newName.length > 24) throw new Error('Слишком длинное имя — до 24 символов');
  if (newName === user.name) return newName;
  return withLock_(() => {
    const list = members_();
    if (list.indexOf(newName) >= 0) throw new Error('Это имя уже занято');
    const props = PropertiesService.getScriptProperties();
    ['items', 'reviews', 'comments'].forEach(key => {
      const col = SCHEMA[key].findIndex(c => c[0] === 'author') + 1;
      const sh = sheet_(key);
      const n = sh.getLastRow() - 1;
      if (n < 1) return;
      const range = sh.getRange(2, col, n, 1);
      range.setValues(range.getValues().map(r => [r[0] === user.name ? newName : r[0]]));
    });
    // Сводка «Купили бы снова» у товаров вида «Имя: Да, Имя: Нет».
    const againCol = SCHEMA.items.findIndex(c => c[0] === 'again') + 1;
    const itemRows = sheet_('items').getLastRow() - 1;
    if (itemRows > 0) {
      const range = sheet_('items').getRange(2, againCol, itemRows, 1);
      range.setValues(range.getValues().map(r => [String(r[0]).split(', ')
        .map(s => (s.indexOf(user.name + ': ') === 0 ? newName + s.slice(user.name.length) : s)).join(', ')]));
    }
    // Ключ входа, подписки на уведомления и отметка «что уже показано» переезжают на новое имя.
    ['KEY_', 'SUBS_', 'SEEN_'].forEach(prefix => {
      const v = props.getProperty(prefix + user.name);
      if (v !== null) props.setProperty(prefix + newName, v);
      props.deleteProperty(prefix + user.name);
    });
    const events = parse_(props.getProperty('EVENTS'), []);
    events.forEach(e => { if (e.author === user.name) e.author = newName; if (e.to === user.name) e.to = newName; });
    props.setProperty('EVENTS', JSON.stringify(events));
    props.setProperty('MEMBERS', JSON.stringify(list.map(n => (n === user.name ? newName : n))));
    user.name = newName;
    return newName;
  });
}

/* ---------- Пуш-уведомления ---------- */

/** Сохраняет подписку этого устройства на уведомления (у человека может быть несколько устройств). */
function savePush(sub) {
  const user = me_();
  if (!sub || !/^https:\/\//.test(sub.endpoint || '')) throw new Error('Неверная подписка');
  return withLock_(() => {
    const props = PropertiesService.getScriptProperties();
    const subs = parse_(props.getProperty('SUBS_' + user.name), []).filter(s => s.endpoint !== sub.endpoint);
    subs.push({ endpoint: sub.endpoint });
    props.setProperty('SUBS_' + user.name, JSON.stringify(subs.slice(-5)));
  });
}

function removePush(endpoint) {
  const user = me_();
  return withLock_(() => removeSub_(user.name, endpoint));
}

/** Отправляет уведомление самому себе — проверить, что всё настроено. */
function testPush() {
  const user = me_();
  const sent = notify_({ type: 'test', author: 'Покупочки', to: user.name, name: 'Уведомления работают 🎉' }, [user.name]);
  if (!sent) throw new Error('На этом аккаунте нет устройств с включёнными уведомлениями');
  return sent;
}

/**
 * Текст для уведомления. Пуш приходит без содержимого (так не нужно шифрование),
 * и service worker на телефоне сразу спрашивает здесь, что именно произошло.
 */
function getNews() {
  const user = me_();
  const props = PropertiesService.getScriptProperties();
  const events = parse_(props.getProperty('EVENTS'), [])
    .filter(e => e.author !== user.name && (!e.to || e.to === user.name));
  const seen = Number(props.getProperty('SEEN_' + user.name) || 0);
  const e = events.find(x => x.t > seen) || events[events.length - 1];
  if (!e) return null;
  props.setProperty('SEEN_' + user.name, String(Math.max(seen, e.t)));
  const what = { item: 'новая покупочка', comment: 'новый комментарий', review: 'оценка вашей покупочки', test: 'проверка' }[e.type] || 'что-то новое';
  const body = e.type === 'comment' ? e.name + ' — «' + e.text + '»' : e.text ? e.name + ' — ' + e.text : e.name;
  return { title: e.author + ': ' + what, body, tag: e.type + '-' + e.t, url: e.itemId ? './?item=' + encodeURIComponent(e.itemId) : './' };
}

/** Запоминает событие и будит устройства остальных участников. Ошибки уведомлений не мешают сохранению. */
function notify_(event, to) {
  try {
    const props = PropertiesService.getScriptProperties();
    const events = parse_(props.getProperty('EVENTS'), []);
    events.push(Object.assign({ t: Date.now() }, event));
    props.setProperty('EVENTS', JSON.stringify(events.slice(-30)));
    return pushTo_(to || members_().filter(n => n !== event.author));
  } catch (e) {
    console.error('Уведомление не отправлено: ' + e);
    if (to) throw e;
    return 0;
  }
}

function pushTo_(names) {
  const props = PropertiesService.getScriptProperties();
  const keys = vapidKeys_();
  const jwts = {};
  const reqs = [], owners = [];
  names.forEach(name => parse_(props.getProperty('SUBS_' + name), []).forEach(sub => {
    const aud = sub.endpoint.match(/^https:\/\/[^/]+/)[0];
    jwts[aud] = jwts[aud] || vapidJwt_(aud, keys.priv);
    reqs.push({ url: sub.endpoint, method: 'post', muteHttpExceptions: true,
      headers: { Authorization: 'vapid t=' + jwts[aud] + ', k=' + keys.pub, TTL: '86400', Urgency: 'high' } });
    owners.push([name, sub.endpoint]);
  }));
  if (!reqs.length) return 0;
  let sent = 0;
  UrlFetchApp.fetchAll(reqs).forEach((res, i) => {
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) sent++;
    else if (code === 404 || code === 410) removeSub_(owners[i][0], owners[i][1]); // подписка больше не действует
    else console.error('Пуш ' + code + ': ' + res.getContentText().slice(0, 200));
  });
  return sent;
}

function removeSub_(name, endpoint) {
  const props = PropertiesService.getScriptProperties();
  const subs = parse_(props.getProperty('SUBS_' + name), []).filter(s => s.endpoint !== endpoint);
  props.setProperty('SUBS_' + name, JSON.stringify(subs));
}

/** Ключи VAPID: создаются один раз и хранятся в свойствах скрипта. Публичный ключ нужен сайту. */
function vapidKeys_() {
  const props = PropertiesService.getScriptProperties();
  let priv = props.getProperty('VAPID_PRIVATE');
  if (!priv) {
    let d = N0;
    while (d === N0) d = P256.mod(bytesToBig_(randomBytes_()), P256.N);
    priv = d.toString(16);
    const Q = P256.mul(d, P256.G);
    props.setProperties({ VAPID_PRIVATE: priv, VAPID_PUBLIC: b64url_([4].concat(bigToBytes_(Q[0]), bigToBytes_(Q[1]))) });
  }
  return { priv, pub: props.getProperty('VAPID_PUBLIC') };
}

/** JWT для заголовка VAPID, подписанный ES256. */
function vapidJwt_(aud, privHex) {
  const enc = obj => b64url_(Utilities.newBlob(JSON.stringify(obj)).getBytes());
  const input = enc({ typ: 'JWT', alg: 'ES256' }) + '.' +
    enc({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: 'https://unseenmoviereviews.github.io/nashi-pokupki/' });
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return input + '.' + b64url_(P256.sign(hash, BigInt('0x' + privHex)));
}

/** Запустите один раз в редакторе: создаст ключи уведомлений и попросит разрешение на отправку запросов. */
function setupPush() {
  Logger.log('Публичный ключ уведомлений: ' + vapidKeys_().pub);
  UrlFetchApp.fetch('https://web.push.apple.com/', { muteHttpExceptions: true });
  Logger.log('Готово: уведомления можно включать в приложении.');
}

// Редактор Apps Script не понимает литералы вида 0n, поэтому малые BigInt заданы так.
const N0 = BigInt(0), N1 = BigInt(1), N2 = BigInt(2), N3 = BigInt(3), N4 = BigInt(4), N8 = BigInt(8);

/*
 * ECDSA на кривой P-256 — в Apps Script её нет, поэтому реализована на BigInt
 * (якобиевы координаты, a = −3). Используется только для подписи VAPID.
 */
const P256 = (() => {
  const p = BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
  const N = BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
  const G = [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
    BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')];
  const mod = (a, m) => { const r = a % m; return r < N0 ? r + m : r; };
  const inv = (a, m) => {
    let r0 = mod(a, m), r1 = m, s0 = N1, s1 = N0;
    while (r1 !== N0) {
      const q = r0 / r1;
      [r0, r1] = [r1, r0 - q * r1];
      [s0, s1] = [s1, s0 - q * s1];
    }
    return mod(s0, m);
  };
  const dbl = P => {
    if (!P || P[1] === N0) return null;
    const [X, Y, Z] = P;
    const delta = mod(Z * Z, p), gamma = mod(Y * Y, p), beta = mod(X * gamma, p);
    const alpha = mod(N3 * (X - delta) * (X + delta), p);
    const X3 = mod(alpha * alpha - N8 * beta, p);
    const Z3 = mod((Y + Z) * (Y + Z) - gamma - delta, p);
    const Y3 = mod(alpha * (N4 * beta - X3) - N8 * gamma * gamma, p);
    return [X3, Y3, Z3];
  };
  const add = (P, Q) => {
    if (!P) return Q;
    if (!Q) return P;
    const [X1, Y1, Z1] = P, [X2, Y2, Z2] = Q;
    const Z1Z1 = mod(Z1 * Z1, p), Z2Z2 = mod(Z2 * Z2, p);
    const U1 = mod(X1 * Z2Z2, p), U2 = mod(X2 * Z1Z1, p);
    const S1 = mod(Y1 * Z2 * Z2Z2, p), S2 = mod(Y2 * Z1 * Z1Z1, p);
    const H = mod(U2 - U1, p), r = mod(S2 - S1, p);
    if (H === N0) return r === N0 ? dbl(P) : null;
    const HH = mod(H * H, p), HHH = mod(H * HH, p), V = mod(U1 * HH, p);
    const X3 = mod(r * r - HHH - N2 * V, p);
    const Y3 = mod(r * (V - X3) - S1 * HHH, p);
    return [X3, Y3, mod(Z1 * Z2 * H, p)];
  };
  // k·(x, y) → аффинная точка [x, y]
  const mul = (k, A) => {
    let R = null;
    const J = [A[0], A[1], N1];
    for (let i = k.toString(2), j = 0; j < i.length; j++) {
      R = dbl(R);
      if (i[j] === '1') R = add(R, J);
    }
    const zi = inv(R[2], p), zi2 = mod(zi * zi, p);
    return [mod(R[0] * zi2, p), mod(R[1] * zi2 * zi, p)];
  };
  // Подпись хеша (32 байта) → r‖s, 64 байта. Нонс «с подстраховкой»: хеш ключа, сообщения и случайных байт.
  const sign = (hash, d) => {
    const e = mod(bytesToBig_(hash), N);
    for (;;) {
      const seed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
        bigToBytes_(d).concat(Array.from(hash), randomBytes_()));
      const k = mod(bytesToBig_(seed), N);
      if (k === N0) continue;
      const r = mod(mul(k, G)[0], N);
      if (r === N0) continue;
      const s = mod(inv(k, N) * (e + r * d), N);
      if (s === N0) continue;
      return bigToBytes_(r).concat(bigToBytes_(s));
    }
  };
  return { N, G, mod, mul, sign };
})();

// Криптостойкие случайные байты: UUID в Apps Script берутся из SecureRandom.
function randomBytes_() {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, [0, 1, 2, 3].map(() => Utilities.getUuid()).join(''));
}

function bytesToBig_(bytes) {
  return BigInt('0x' + Array.from(bytes, b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('') || '0');
}

function bigToBytes_(n) {
  const hex = n.toString(16).padStart(64, '0');
  const out = [];
  for (let i = 0; i < 64; i += 2) out.push(parseInt(hex.slice(i, i + 2), 16));
  return out;
}

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes.map(b => (b > 127 ? b - 256 : b))).replace(/=+$/, '');
}

/* ---------- Вспомогательные функции ---------- */

function me_() {
  if (!currentUser) throw new Error('Нет доступа');
  return currentUser;
}

/** Текущий список имён: после переименований он хранится в свойствах скрипта. */
function members_() {
  const saved = PropertiesService.getScriptProperties().getProperty('MEMBERS');
  return saved ? JSON.parse(saved) : MEMBERS.slice();
}

function userByKey_(key) {
  if (!key) return null;
  const props = PropertiesService.getScriptProperties();
  const name = members_().find(n => props.getProperty('KEY_' + n) === String(key));
  return name ? { name } : null;
}

let ss__ = null;
function ss_() {
  return ss__ || (ss__ = SpreadsheetApp.openById(prop_('SPREADSHEET_ID')));
}

function prop_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error('Приложение не настроено: запустите setup() в редакторе скриптов.');
  return v;
}

function sheet_(key) {
  return ss_().getSheetByName(SHEETS[key]);
}

function read_(key) {
  const cols = SCHEMA[key].map(c => c[0]);
  const sh = sheet_(key);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
    return o;
  });
}

function write_(key, o) {
  const row = SCHEMA[key].map(c => cell_(o[c[0]]));
  const sh = sheet_(key);
  // Столбцы, добавленные в схему после создания листа, получают заголовок при первой записи.
  const had = sh.getLastColumn();
  if (had < row.length) {
    sh.getRange(1, 1, 1, row.length).setValues([SCHEMA[key].map(c => c[1])]).setFontWeight('bold');
    SCHEMA[key].forEach((c, i) => { if (i >= had && c[1][0] === '_') sh.hideColumns(i + 1); });
  }
  if (o._row) sh.getRange(o._row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

// Текст, начинающийся с = + - @, таблица приняла бы за формулу.
function cell_(v) {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function deleteRows_(key, pred) {
  const sh = sheet_(key);
  read_(key).filter(pred).map(r => r._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

function trashPhoto_(p) {
  [p.full, p.thumb].forEach(id => {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* файл уже удалён */ }
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function parse_(s, fallback) {
  try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; }
}

function avg_(vals) {
  return vals.length ? Math.round(vals.reduce((a, b) => a + b, 0) / vals.length * 10) / 10 : '';
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}
