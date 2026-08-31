/**
 * Трекер расходов — backend на Google Apps Script.
 * Работает поверх Google Таблицы (лист «Транзакции») и служит API для index.html,
 * вебхука Точка.API и парсера писем Озон Банка.
 *
 * НАСТРОЙКА ПЕРЕД ИСПОЛЬЗОВАНИЕМ:
 * 1. Замените SHARED_SECRET ниже на собственный длинный случайный пароль.
 *    Тот же пароль впишите в дашборде (шестерёнка → Пароль) после публикации на GitHub Pages.
 * 2. Deploy → Manage deployments → редактируйте СУЩЕСТВУЮЩЕЕ развёртывание (не создавайте новое),
 *    Execute as: Me, Who has access: Anyone.
 * 3. В Триггерах (часики слева) добавьте time-driven trigger на checkGmailForTransactions (каждые 10 минут).
 */

const SHARED_SECRET = '_0uUtUTbXSp3dDzeRI9Bx6gRCnaI5YKs';
const SHEET_NAME = 'Транзакции';
const HEADERS = ['id', 'Добавлено', 'Дата операции', 'Сумма', 'Описание', 'Источник', 'Категория', 'Метка', 'ФайлID'];
const DDS_SHEET_NAME = 'ДДС';

// Настройки поиска писем Озон Банка — ЧЕРНОВЫЕ, нужно уточнить по реальному письму.
const OZON_SENDER_QUERY = 'from:(noreply@ozon.ru OR notify@ozon.ru) subject:(операция OR списание OR платёж)';
const OZON_LABEL_PROCESSED = 'ozon-tracker-processed';

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowsToObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .filter(r => r[0] !== '' && r[0] != null)
    .map(r => {
      const obj = {};
      HEADERS.forEach((h, i) => { obj[h] = r[i]; });
      return {
        id: String(obj['id']),
        date: formatDate_(obj['Дата операции']),
        desc: obj['Описание'] || '',
        amount: Number(obj['Сумма']) || 0,
        source: obj['Источник'] || '',
        category: obj['Категория'] || '',
        type: obj['Метка'] || 'unset',
        fileId: obj['ФайлID'] || '',
      };
    });
}

function formatDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return s.slice(0, 10);
}

function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // sheet row number
  }
  return -1;
}

function appendTxn_(sheet, txn) {
  const id = txn.id || Utilities.getUuid();
  sheet.appendRow([
    id,
    new Date(),
    txn.date || '',
    Number(txn.amount) || 0,
    txn.desc || '',
    txn.source || '',
    txn.category || 'other',
    txn.type || 'unset',
    txn.fileId || '',
  ]);
  return id;
}

function appendTxnsBatch_(sheet, txns) {
  // Один вызов setValues() вместо appendRow() в цикле — для импорта выписки
  // на десятки строк это секунды вместо минут и убирает гонку с параллельным
  // запросом "Обновить" с сайта, который может застать ещё не дописанные данные.
  if (!txns || !txns.length) return [];
  const now = new Date();
  const ids = [];
  const rows = txns.map(function (txn) {
    const id = txn.id || Utilities.getUuid();
    ids.push(id);
    return [
      id,
      now,
      txn.date || '',
      Number(txn.amount) || 0,
      txn.desc || '',
      txn.source || '',
      txn.category || 'other',
      txn.type || 'unset',
      txn.fileId || '',
    ];
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
  return ids;
}

// ---------------------------------------------------------------------------
// HTTP entry points
// ---------------------------------------------------------------------------

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.secret !== SHARED_SECRET) {
    return jsonOut_({ ok: false, error: 'unauthorized' });
  }
  const sheet = getSheet_();
  const txns = rowsToObjects_(sheet);
  return jsonOut_({ ok: true, txns: txns });
}

function doPost(e) {
  const rawBody = (e && e.postData && e.postData.contents) || '';
  let body = null;
  try {
    body = JSON.parse(rawBody);
  } catch (err) {
    // Не JSON — это вебхук Точки: там в теле "голая" строка JWT (Content-Type: text/plain),
    // см. handleTochkaWebhook_.
    body = null;
  }

  // 1) Запросы с самого сайта — у них всегда явное поле action + правильный secret.
  if (body && body.action) {
    if (body.secret !== SHARED_SECRET) return jsonOut_({ ok: false, error: 'unauthorized' });
    return handleSiteAction_(body);
  }

  // 2) «Быстрая запись» — для будущего шортката iPhone: src=quick или просто secret без action.
  if (body && (body.src === 'quick' || (body.secret && !body.action))) {
    if (body.secret !== SHARED_SECRET) return jsonOut_({ ok: false, error: 'unauthorized' });
    return handleQuickAdd_(body);
  }

  // 3) Иначе считаем это вебхуком Точка.API — тело не JSON, это подписанная JWT-строка.
  return handleTochkaWebhook_(rawBody);
}

function handleSiteAction_(body) {
  const sheet = getSheet_();
  switch (body.action) {
    case 'add': {
      const id = appendTxn_(sheet, body.txn || {});
      rebuildDDS_();
      return jsonOut_({ ok: true, id: id });
    }
    case 'bulk_add': {
      const txns = body.txns || [];
      const ids = appendTxnsBatch_(sheet, txns);
      rebuildDDS_();
      return jsonOut_({ ok: true, ids: ids });
    }
    case 'update_fields': {
      const row = findRowById_(sheet, body.id);
      if (row < 0) return jsonOut_({ ok: false, error: 'not_found' });
      const fields = body.fields || {};
      if (fields.category !== undefined) sheet.getRange(row, 7).setValue(fields.category);
      if (fields.type !== undefined) sheet.getRange(row, 8).setValue(fields.type);
      if (fields.desc !== undefined) sheet.getRange(row, 5).setValue(fields.desc);
      if (fields.date !== undefined) sheet.getRange(row, 3).setValue(fields.date);
      if (fields.amount !== undefined) sheet.getRange(row, 4).setValue(fields.amount);
      if (fields.source !== undefined) sheet.getRange(row, 6).setValue(fields.source);
      rebuildDDS_();
      return jsonOut_({ ok: true });
    }
    case 'delete_batch': {
      const ids = body.ids || [];
      deleteRowsByIds_(sheet, ids);
      rebuildDDS_();
      return jsonOut_({ ok: true, deleted: ids.length });
    }
    case 'delete_all': {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      rebuildDDS_();
      return jsonOut_({ ok: true });
    }
    default:
      return jsonOut_({ ok: false, error: 'unknown_action' });
  }
}

function deleteRowsByIds_(sheet, ids) {
  const idSet = new Set(ids.map(String));
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Удаляем снизу вверх, чтобы номера строк не сбивались.
  for (let i = idValues.length - 1; i >= 0; i--) {
    if (idSet.has(String(idValues[i][0]))) {
      sheet.deleteRow(i + 2);
    }
  }
}

function handleQuickAdd_(body) {
  const sheet = getSheet_();
  const id = appendTxn_(sheet, {
    date: body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    desc: body.desc || '(быстрая запись)',
    amount: -Math.abs(Number(body.amount) || 0),
    source: body.source || 'Т-Банк',
    category: body.category || guessCategory_(body.desc || ''),
    type: body.type || 'unset',
  });
  rebuildDDS_();
  return jsonOut_({ ok: true, id: id });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Точка.API — вебхук по операциям счёта ИП
// ---------------------------------------------------------------------------
// Тело запроса — не JSON, а «голая» подписанная JWT-строка (Content-Type: text/plain).
// Подпись (RS256) здесь НЕ проверяется — сознательное упрощение для личного одиночного
// использования (публичный ключ Точки: https://enter.tochka.com/doc/openapi/static/keys/public).
// Если это станет важно (например, эндпоинт станет достижим извне не только для Точки),
// нужно добавить проверку подписи перед доверием содержимому.
// Форма самого payload внутри JWT ещё не подтверждена реальным вебхуком — events/operation
// ниже это предположение по документации; при первом реальном вызове смотрите Executions в
// Apps Script (console.log payload) и поправьте разбор под то, что реально пришло.

function handleTochkaWebhook_(rawBody) {
  let payload = null;
  try {
    payload = decodeJwtPayload_(rawBody);
  } catch (err) {
    console.error('Tochka webhook: не смог декодировать JWT. err=' + err + ' rawBody=' + rawBody);
  }
  if (!payload) {
    // Точка проверяет доступность URL тестовым вызовом при создании/редактировании вебхука —
    // отвечаем 200 даже если не смогли разобрать тело, иначе вебхук не создастся.
    console.error('Tochka webhook: пустой/неразбираемый payload. rawBody=' + rawBody);
    return jsonOut_({ ok: true, note: 'unparsed' });
  }
  console.log('Tochka webhook payload: ' + JSON.stringify(payload));
  try {
    const events = payload.events || (payload.Data && payload.Data.events) || [payload];
    const sheet = getSheet_();
    const ids = [];
    events.forEach(function (evt) {
      const op = evt.operation || evt.Data || evt;
      const parsed = parseTochka_(op);
      if (parsed) ids.push(appendTxn_(sheet, parsed));
    });
    if (ids.length) rebuildDDS_();
    return jsonOut_({ ok: true, ids: ids });
  } catch (err) {
    console.error('Tochka webhook handling error: ' + err + ' payload=' + JSON.stringify(payload));
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function decodeJwtPayload_(jwt) {
  const token = (jwt || '').trim();
  const parts = token.split('.');
  if (parts.length < 2) return null;
  let b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  while (b64.length % 4) b64 += '=';
  const bytes = Utilities.base64Decode(b64);
  const json = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  return JSON.parse(json);
}

function parseTochka_(op) {
  // ЧЕРНОВИК — названия полей нужно свериться с реальным вебхуком Точки после подключения JWT-токена.
  if (!op) return null;
  const amountRaw = op.amount || op.payment_amount || (op.amount_details && op.amount_details.amount);
  if (amountRaw == null) return null;
  const amount = -Math.abs(Number(amountRaw)); // Точка обычно шлёт списания положительным числом
  const desc = op.purpose || op.description || op.payer_name || '(операция Точка)';
  const date = op.date || op.operation_date || op.created_at || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {
    id: op.operation_id || op.id || Utilities.getUuid(),
    date: formatDate_(date),
    desc: desc,
    amount: amount,
    source: 'Точка',
    category: guessCategory_(desc),
    type: 'unset',
  };
}

// --- Регистрация вебхука в Точка.API ---------------------------------------
// Ключ и client_id берутся из Script Properties (Project Settings → Script Properties),
// а не из кода — чтобы они не попадали в git и в переписку.
// По документации Точки ключ используется напрямую как Bearer-токен, без обмена (нет refresh —
// когда истечёт срок действия, нужно перевыпустить ключ в приложении банка и обновить property).

const TOCHKA_BASE_URL_ = 'https://enter.tochka.com/uapi';
const TOCHKA_WEBHOOK_EVENTS_ = ['incomingPayment', 'outgoingPayment', 'incomingSbpPayment', 'incomingSbpB2BPayment', 'acquiringInternetPayment'];

function tochkaProp_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) throw new Error('Нет ' + name + ' в Script Properties (Project Settings → Script Properties)');
  return value;
}

// Запустите вручную (▶ в редакторе) один раз, чтобы создать вебхук на текущий деплой.
function registerTochkaWebhook() {
  const clientId = tochkaProp_('TOCHKA_CLIENT_ID');
  const jwt = tochkaProp_('TOCHKA_JWT_KEY');
  const webAppUrl = ScriptApp.getService().getUrl();
  const resp = UrlFetchApp.fetch(TOCHKA_BASE_URL_ + '/webhook/v1.0/' + encodeURIComponent(clientId), {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + jwt },
    payload: JSON.stringify({ webhooksList: TOCHKA_WEBHOOK_EVENTS_, url: webAppUrl }),
    muteHttpExceptions: true,
  });
  Logger.log('registerTochkaWebhook -> URL=%s code=%s body=%s', webAppUrl, resp.getResponseCode(), resp.getContentText());
  return resp.getContentText();
}

// Проверить, что сейчас зарегистрировано (GET).
function checkTochkaWebhooks() {
  const clientId = tochkaProp_('TOCHKA_CLIENT_ID');
  const jwt = tochkaProp_('TOCHKA_JWT_KEY');
  const resp = UrlFetchApp.fetch(TOCHKA_BASE_URL_ + '/webhook/v1.0/' + encodeURIComponent(clientId), {
    method: 'get',
    headers: { Authorization: 'Bearer ' + jwt },
    muteHttpExceptions: true,
  });
  Logger.log('checkTochkaWebhooks -> code=%s body=%s', resp.getResponseCode(), resp.getContentText());
  return resp.getContentText();
}

// ---------------------------------------------------------------------------
// Озон Банк — разбор писем через Gmail (запускается по таймеру каждые 10 минут)
// ---------------------------------------------------------------------------

function checkGmailForTransactions() {
  const threads = GmailApp.search(OZON_SENDER_QUERY + ' -label:' + OZON_LABEL_PROCESSED, 0, 20);
  if (!threads.length) return;

  let label = GmailApp.getUserLabelByName(OZON_LABEL_PROCESSED);
  if (!label) label = GmailApp.createLabel(OZON_LABEL_PROCESSED);

  const sheet = getSheet_();
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      const parsed = parseOzonEmail_(msg.getPlainBody() || msg.getBody());
      if (parsed) {
        parsed.fileId = msg.getId();
        appendTxn_(sheet, parsed);
      }
    });
    thread.addLabel(label);
  });
  rebuildDDS_();
}

function parseOzonEmail_(text) {
  // ЧЕРНОВИК — регулярку нужно поправить по реальному письму Озон Банка (структура письма пока не подтверждена).
  if (!text) return null;
  const amountMatch = text.match(/(?:списани[ея]|оплата|покупка)[^\d]*(\d[\d\s]*[.,]?\d{0,2})\s*(?:₽|руб)/i);
  if (!amountMatch) return null;
  const amount = -Math.abs(parseFloat(amountMatch[1].replace(/\s/g, '').replace(',', '.')));

  const descMatch = text.match(/(?:в|у|получатель)[:\s]+([^\n\r]{2,60})/i);
  const desc = descMatch ? descMatch[1].trim() : '(операция Озон Банк)';

  const dateMatch = text.match(/(\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4})/);
  const date = dateMatch ? normalizeDate_(dateMatch[1]) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  return {
    id: Utilities.getUuid(),
    date: date,
    desc: desc,
    amount: amount,
    source: 'Озон Банк',
    category: guessCategory_(desc),
    type: 'unset',
  };
}

function normalizeDate_(s) {
  const m = s.match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{2,4})/);
  if (!m) return s;
  const y = m[3].length === 2 ? '20' + m[3] : m[3];
  return y + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
}

// ---------------------------------------------------------------------------
// Простая авто-категоризация по ключевым словам (совпадает с логикой в index.html)
// ---------------------------------------------------------------------------

const CATEGORY_RULES = [
  { category: 'groceries', keys: ['пятерочка', 'перекрест', 'магнит', 'ашан', 'лента', 'вкусвилл', 'дикси', 'окей', 'spar', 'metro'] },
  { category: 'cafe', keys: ['кафе', 'ресторан', 'старбакс', 'кофе', 'coffee', 'kfc', 'макдон', 'mcdonald', 'burger', 'пицц', 'sushi', 'суши'] },
  { category: 'transport', keys: ['такси', 'taxi', 'yandex go', 'uber', 'ржд', 'aeroflot', 'авиа', 'азс', 'лукойл', 'газпромнефть', 'shell', 'каршер', 'метрополитен'] },
  { category: 'housing', keys: ['жкх', 'коммунал', 'электросет', 'водоканал', 'управляющая компания', 'квартплата'] },
  { category: 'comms', keys: ['мтс', 'билайн', 'мегафон', 'tele2', 'теле2', 'ростелеком'] },
  { category: 'health', keys: ['аптека', 'клиника', 'стоматолог', 'больниц', 'медцентр'] },
  { category: 'shopping', keys: ['zara', 'h&m', 'lamoda', 'ecco', 'салон красот', 'парикмахер', 'косметик'] },
  { category: 'marketplace', keys: ['wildberries', 'ozon.ru', 'яндекс маркет', 'aliexpress', 'avito'] },
  { category: 'fun', keys: ['кино', 'cinema', 'netflix', 'spotify', 'steam', 'ivi', 'okko', 'театр', 'музей'] },
  { category: 'subs', keys: ['подписк', 'apple.com/bill', 'google play', 'яндекс плюс', 'icloud'] },
  { category: 'transfers', keys: ['перевод', 'снятие наличных', 'атм', 'atm', 'cash withdraw'] },
];

function guessCategory_(desc) {
  const d = String(desc || '').toLowerCase();
  for (const rule of CATEGORY_RULES) {
    for (const k of rule.keys) {
      if (d.indexOf(k) !== -1) return rule.category;
    }
  }
  return 'other';
}

// ---------------------------------------------------------------------------
// ДДС — автоматически оформленный отчёт на отдельном листе
// ---------------------------------------------------------------------------
// Строится заново при каждой записи в «Транзакции» (см. вызовы rebuildDDS_
// в handleSiteAction_/handleQuickAdd_/handleTochkaWebhook_/checkGmailForTransactions).
// Обёрнуто в try/catch, чтобы сбой форматирования отчёта никогда не ронял
// сам факт сохранения операции.

const CATEGORY_LABELS_ = {
  groceries: 'Продукты', cafe: 'Кафе и рестораны', transport: 'Транспорт',
  housing: 'Жильё и ЖКХ', comms: 'Связь и интернет', health: 'Здоровье и аптеки',
  shopping: 'Одежда и красота', marketplace: 'Маркетплейсы', fun: 'Развлечения и хобби',
  subs: 'Подписки и сервисы', transfers: 'Переводы и снятия', other: 'Прочее',
};
const CATEGORY_ORDER_ = ['groceries', 'cafe', 'transport', 'housing', 'comms', 'health',
  'shopping', 'marketplace', 'fun', 'subs', 'transfers', 'other'];
const RU_MONTHS_SHORT_ = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];

function rebuildDDS_() {
  try {
    rebuildDDSUnsafe_();
  } catch (err) {
    console.error('rebuildDDS_ failed: ' + err);
  }
}

// Обёртка без подчёркивания в конце имени — Apps Script прячет из выпадающего
// списка "Выполнить" все функции вида *_, поэтому для ручного запуска/проверки
// в редакторе используйте именно эту функцию.
function manualRebuildDDS() {
  rebuildDDS_();
}

function rebuildDDSUnsafe_() {
  const txns = rowsToObjects_(getSheet_()).filter(function (t) { return t.amount < 0; });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DDS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DDS_SHEET_NAME);
  } else {
    sheet.clear();
    sheet.clearFormats();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
  }

  if (!txns.length) {
    sheet.getRange(1, 1).setValue('Нет операций для отчёта — добавьте хотя бы одну транзакцию.');
    return;
  }

  const monthsSet = {};
  txns.forEach(function (t) { if (t.date) monthsSet[t.date.slice(0, 7)] = true; });
  const months = Object.keys(monthsSet).sort();
  const monthLabel = function (m) {
    const parts = m.split('-');
    return RU_MONTHS_SHORT_[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
  };

  const HEADER_BG = '#1f2937';
  const HEADER_FG = '#ffffff';
  const SUBHEAD_BG = '#e5e7eb';
  const TOTAL_BG = '#f3f4f6';
  const BORDER_COLOR = '#d1d5db';
  const NUM_FMT = '#,##0 ₽';

  let row = 1;

  // ---- Заголовок ----
  const titleColSpan = months.length + 2;
  sheet.getRange(row, 1, 1, titleColSpan).merge()
    .setValue('ДДС — движение денежных средств')
    .setFontSize(14).setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);
  row += 1;
  sheet.getRange(row, 1).setValue('Сформировано автоматически: ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'));
  sheet.getRange(row, 1).setFontColor('#6b7280').setFontSize(9).setFontStyle('italic');
  row += 2;

  // ---- Таблица: категории × месяцы ----
  const catTableTop = row;
  const catHeader = ['Категория'].concat(months.map(monthLabel)).concat(['Итого']);
  sheet.getRange(row, 1, 1, catHeader.length).setValues([catHeader])
    .setFontWeight('bold').setBackground(SUBHEAD_BG);
  row += 1;

  const catTotalsByMonth = {};
  months.forEach(function (m) { catTotalsByMonth[m] = 0; });
  let grandTotal = 0;
  const catRows = [];

  CATEGORY_ORDER_.forEach(function (catId) {
    let catSum = 0;
    const vals = months.map(function (m) {
      const sum = txns
        .filter(function (t) { return t.category === catId && t.date && t.date.slice(0, 7) === m; })
        .reduce(function (s, t) { return s + (-t.amount); }, 0);
      catTotalsByMonth[m] += sum;
      catSum += sum;
      return sum || 0;
    });
    if (catSum > 0) {
      grandTotal += catSum;
      catRows.push([CATEGORY_LABELS_[catId] || catId].concat(vals).concat([catSum]));
    }
  });

  if (catRows.length) {
    sheet.getRange(row, 1, catRows.length, catHeader.length).setValues(catRows);
    row += catRows.length;
  }

  const totalRow = ['Итого'].concat(months.map(function (m) { return catTotalsByMonth[m] || 0; })).concat([grandTotal]);
  sheet.getRange(row, 1, 1, totalRow.length).setValues([totalRow])
    .setFontWeight('bold').setBackground(TOTAL_BG);
  const catTableBottom = row;
  row += 2;

  sheet.getRange(catTableTop + 1, 2, catTableBottom - catTableTop, months.length + 1).setNumberFormat(NUM_FMT);
  sheet.getRange(catTableTop, 1, catTableBottom - catTableTop + 1, catHeader.length)
    .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

  // ---- По источникам (банкам) ----
  const bankTitleRow = row;
  sheet.getRange(row, 1).setValue('По источникам').setFontWeight('bold').setFontSize(12);
  row += 1;
  const byBank = {};
  txns.forEach(function (t) { const k = t.source || 'Другое'; byBank[k] = (byBank[k] || 0) + (-t.amount); });
  const bankEntries = Object.entries(byBank).sort(function (a, b) { return b[1] - a[1]; });
  sheet.getRange(row, 1, 1, 2).setValues([['Источник', 'Сумма']]).setFontWeight('bold').setBackground(SUBHEAD_BG);
  row += 1;
  const bankDataTop = row;
  bankEntries.forEach(function (entry) {
    sheet.getRange(row, 1, 1, 2).setValues([[entry[0], entry[1]]]);
    row += 1;
  });
  if (bankEntries.length) {
    sheet.getRange(bankDataTop, 2, bankEntries.length, 1).setNumberFormat(NUM_FMT);
    sheet.getRange(bankTitleRow + 1, 1, row - bankTitleRow - 1, 2)
      .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  }
  row += 2;

  // ---- Личное / Рабочее / Без метки ----
  sheet.getRange(row, 1).setValue('Личное vs рабочее').setFontWeight('bold').setFontSize(12);
  row += 1;
  const byType = { personal: 0, work: 0, unset: 0 };
  txns.forEach(function (t) {
    const k = t.type === 'personal' ? 'personal' : (t.type === 'work' ? 'work' : 'unset');
    byType[k] += (-t.amount);
  });
  const typeLabels = { personal: 'Личное', work: 'Рабочее', unset: 'Без метки' };
  sheet.getRange(row, 1, 1, 2).setValues([['Метка', 'Сумма']]).setFontWeight('bold').setBackground(SUBHEAD_BG);
  row += 1;
  const typeDataTop = row;
  ['personal', 'work', 'unset'].forEach(function (k) {
    sheet.getRange(row, 1, 1, 2).setValues([[typeLabels[k], byType[k]]]);
    row += 1;
  });
  sheet.getRange(typeDataTop, 2, 3, 1).setNumberFormat(NUM_FMT);
  sheet.getRange(typeDataTop - 1, 1, 4, 2)
    .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

  // ---- Общее оформление листа ----
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, catHeader.length);
  sheet.setColumnWidth(1, 190);
}
