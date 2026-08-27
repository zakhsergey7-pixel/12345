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

const SHARED_SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_ПАРОЛЬ';
const SHEET_NAME = 'Транзакции';
const HEADERS = ['id', 'Добавлено', 'Дата операции', 'Сумма', 'Описание', 'Источник', 'Категория', 'Метка', 'ФайлID'];

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
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    // Не JSON — может быть form-encoded вебхук Точки, пробуем распарсить как есть.
    body = (e && e.parameter) || {};
  }

  // 1) Запросы с самого сайта — у них всегда явное поле action + правильный secret.
  if (body.action) {
    if (body.secret !== SHARED_SECRET) return jsonOut_({ ok: false, error: 'unauthorized' });
    return handleSiteAction_(body);
  }

  // 2) «Быстрая запись» — для будущего шортката iPhone: src=quick или просто secret без action.
  if (body.src === 'quick' || (body.secret && !body.action)) {
    if (body.secret !== SHARED_SECRET) return jsonOut_({ ok: false, error: 'unauthorized' });
    return handleQuickAdd_(body);
  }

  // 3) Иначе считаем это вебхуком Точка.API (у него своя структура тела, без нашего secret/action).
  return handleTochkaWebhook_(body);
}

function handleSiteAction_(body) {
  const sheet = getSheet_();
  switch (body.action) {
    case 'add': {
      const id = appendTxn_(sheet, body.txn || {});
      return jsonOut_({ ok: true, id: id });
    }
    case 'bulk_add': {
      const txns = body.txns || [];
      const ids = txns.map(t => appendTxn_(sheet, t));
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
      return jsonOut_({ ok: true });
    }
    case 'delete_batch': {
      const ids = body.ids || [];
      deleteRowsByIds_(sheet, ids);
      return jsonOut_({ ok: true, deleted: ids.length });
    }
    case 'delete_all': {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
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
  return jsonOut_({ ok: true, id: id });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Точка.API — вебхук по операциям счёта ИП
// ---------------------------------------------------------------------------
// ПРЕДПОЛОЖЕНИЕ по документации Точка.API, нужно свериться с реальным payload и поправить пути ниже.
// Обычно Точка присылает объект вида { events: [ { type: 'acquiring'|'operation', operation: {...} } ] }
// либо одиночную операцию верхнего уровня — обработайте оба варианта при отладке через logEvent.

function handleTochkaWebhook_(body) {
  try {
    const events = body.events || [body];
    const sheet = getSheet_();
    const ids = [];
    events.forEach(function (evt) {
      const op = evt.operation || evt;
      const parsed = parseTochka_(op);
      if (parsed) ids.push(appendTxn_(sheet, parsed));
    });
    return jsonOut_({ ok: true, ids: ids });
  } catch (err) {
    // Логируем сырое тело, чтобы можно было свериться с реальным форматом Точки и поправить parseTochka_.
    console.error('Tochka webhook parse error: ' + err + ' body=' + JSON.stringify(body));
    return jsonOut_({ ok: false, error: String(err) });
  }
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
