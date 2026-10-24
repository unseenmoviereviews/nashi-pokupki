/**
 * Наши покупки — серверная часть (Google Apps Script).
 * Каталог того, что мы уже купили, с оценками по критериям.
 * Сайт лежит на GitHub Pages и обращается сюда как к API. Вход — по личному ключу,
 * ключи хранятся только в свойствах скрипта (в коде и на GitHub их нет).
 *
 * 1. Впишите имена в MEMBERS и (по желанию) почту пары в SHARE_WITH.
 * 2. Один раз запустите функцию setup() — в журнале появятся ключи входа.
 * 3. Разверните как веб-приложение (см. README.md).
 */
const MEMBERS = ['Я', 'Солнышко'];

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
    ['avg', 'Средняя оценка'], ['photoLinks', 'Фото'], ['photos', '_photos']],
  reviews: [['itemId', 'ID товара'], ['itemName', 'Товар'], ['author', 'Автор'], ['updated', 'Обновлено'],
    ['avg', 'Средняя'], ['ratingsText', 'Оценки'], ['ratings', '_ratings']],
  comments: [['id', 'ID'], ['itemId', 'ID товара'], ['itemName', 'Товар'], ['author', 'Автор'],
    ['date', 'Дата'], ['text', 'Текст']],
  criteria: [['category', 'Категория'], ['list', 'Критерии (через запятую)']],
};

/* ---------- API ---------- */

// Функции, которые может вызывать сайт.
const API = { getData, saveItem, deleteItem, saveReview, addComment, deleteComment, uploadPhoto, getThumbs, getPhoto };
let currentUser = null;

function doGet() {
  return ContentService.createTextOutput('API «Наши покупки» работает.');
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
  MEMBERS.forEach(name => {
    if (!props.getProperty('KEY_' + name)) props.setProperty('KEY_' + name, Utilities.getUuid().replace(/-/g, ''));
  });
  showKeys();
}

/** Печатает ключи входа в журнал выполнения. */
function showKeys() {
  const props = PropertiesService.getScriptProperties();
  MEMBERS.forEach(name => Logger.log(name + ': ' + props.getProperty('KEY_' + name)));
}

/** Если ключ утёк: запустите, затем showKeys() и войдите заново с новыми ключами. */
function resetKeys() {
  const props = PropertiesService.getScriptProperties();
  MEMBERS.forEach(name => props.setProperty('KEY_' + name, Utilities.getUuid().replace(/-/g, '')));
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
      description: it.description, photos: parse_(it.photos, []), reviews: [], comments: [],
    };
    byId[it.id] = out;
    return out;
  });
  read_('reviews').forEach(r => {
    if (byId[r.itemId]) byId[r.itemId].reviews.push({ author: r.author, updated: r.updated, ratings: parse_(r.ratings, {}) });
  });
  read_('comments').forEach(c => {
    if (byId[c.itemId]) byId[c.itemId].comments.push({ id: c.id, author: c.author, date: c.date, text: c.text });
  });
  return { me: user.name, members: MEMBERS, criteria, again: AGAIN, items };
}

function saveItem(d) {
  const user = me_();
  return withLock_(() => {
    let it = null;
    if (d.id) {
      it = read_('items').find(x => x.id === d.id);
      if (!it) throw new Error('Товар не найден — возможно, его удалили.');
      it.created = new Date(it.created);
    } else {
      it = { id: newId_(), created: new Date(), author: user.name, avg: '' };
    }
    ['name', 'url', 'shop', 'currency', 'category', 'again', 'description'].forEach(k => {
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
    write_('items', it);
    return it.id;
  });
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

/** Сохраняет оценку текущего пользователя (по одной на человека на товар). */
function saveReview(itemId, ratings) {
  const user = me_();
  return withLock_(() => {
    const it = read_('items').find(x => x.id === itemId);
    if (!it) throw new Error('Товар не найден');
    const clean = {};
    Object.keys(ratings || {}).forEach(k => {
      const v = Math.round(Number(ratings[k]));
      if (v >= 1 && v <= 5) clean[String(k)] = v;
    });
    const reviews = read_('reviews').filter(r => r.itemId === itemId);
    const mine = reviews.find(r => r.author === user.name);
    if (Object.keys(clean).length) {
      write_('reviews', {
        _row: mine && mine._row, itemId, itemName: it.name, author: user.name, updated: new Date(),
        avg: avg_(Object.values(clean)), ratingsText: Object.keys(clean).map(k => k + ': ' + clean[k]).join(', '),
        ratings: JSON.stringify(clean),
      });
    } else if (mine) {
      sheet_('reviews').deleteRow(mine._row);
    }
    // Пересчитываем общую среднюю товара по оценкам всех участников.
    const all = [];
    reviews.filter(r => r.author !== user.name).forEach(r => all.push(...Object.values(parse_(r.ratings, {}))));
    all.push(...Object.values(clean));
    it.avg = avg_(all);
    it.created = new Date(it.created);
    if (it.bought) it.bought = new Date(it.bought);
    write_('items', it);
  });
}

function addComment(itemId, text) {
  const user = me_();
  text = String(text || '').trim();
  if (!text) return;
  return withLock_(() => {
    const it = read_('items').find(x => x.id === itemId);
    if (!it) throw new Error('Товар не найден');
    write_('comments', { id: newId_(), itemId, itemName: it.name, author: user.name, date: new Date(), text });
  });
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

/* ---------- Вспомогательные функции ---------- */

function me_() {
  if (!currentUser) throw new Error('Нет доступа');
  return currentUser;
}

function userByKey_(key) {
  if (!key) return null;
  const props = PropertiesService.getScriptProperties();
  const name = MEMBERS.find(n => props.getProperty('KEY_' + n) === String(key));
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
