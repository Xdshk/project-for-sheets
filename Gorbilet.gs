function processGorbiletOrders() {
  const SHEET_NAME = "Автомат";
  const SEARCH_QUERY = 'from:Горбилет subject:("Оплачен заказ") newer_than:1d';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const logSheet = ss.getSheetByName("Errors") || ss.insertSheet("Errors");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const existingKeys = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastCol)
        .getValues()
        .map(r => normalizeKey(`${r[0] || ""}_${r[11] || ""}_${r[5] || ""}_${r[2] || ""}_${r[16] || ""}`))
    : [];

  const threads = GmailApp.search(SEARCH_QUERY);
  const now = new Date();
  const DAY = 24 * 60 * 60 * 1000;

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      try {
        if (now - message.getDate() > DAY) continue;
        const body = message.getPlainBody();
        const orders = parseGorbiletEmail(body);

        console.log("Горбилет | Письмо " + message.getId() + ": найдено заказов - " + orders.length);

        for (const order of orders) {
          const uniqueKey = normalizeKey(`${order.orderId}_${order.productName}_${order.quantity}_${order.firstName}_${order.ticketType}`);
          if (existingKeys.includes(uniqueKey)) {
            console.log("Пропущен дубликат: " + uniqueKey);
            continue;
          }

          sheet.appendRow([
            order.orderId,      // 0
            order.orderDate,    // 1
            order.firstName,    // 2
            order.lastName,     // 3
            "",                 // 4
            order.quantity,     // 5
            "",                 // 6
            order.phone,        // 7
            "",                 // 8
            order.email,        // 9
            "",                 // 10
            order.productName,  // 11
            order.tourDate,     // 12
            order.quantity,     // 13
            "'" + order.totalAmount, // 14
            "Горбилет",         // 15 (Источник)
            order.ticketType,   // 16
            order.note || "",   // 17
            ""                  // 18
          ]);

          existingKeys.push(uniqueKey);
        }
      } catch (e) {
        console.error("Ошибка в письме " + message.getId() + ": " + e.message);
        if (typeof logError !== 'undefined') logError(logSheet, e, message.getId());
      }
    }
  }
}

function parseGorbiletEmail(body) {
  const results = [];
  if (!body || typeof body !== "string") return results;

  try {
    const orderId = safeMatch(body, /Заказ\s*№\s*(\d+)/i, 1) || "";
    const nameMatch = safeMatch(body, /Имя покупателя:\s*([^\r\n]+)/i, 1) || "";
    const parts = nameMatch.trim() ? nameMatch.trim().split(/\s+/) : [];
    const firstName = parts.length ? parts[0] : "";
    const lastName = parts.length > 1 ? parts.slice(1).join(" ") : "";

    const phone = normalizePhone(safeMatch(body, /Телефон:\s*(\+?\d[\d\-\s\(\)]*)/i, 1) || "");
    const email = safeMatch(body, /Email:\s*([\w.\-+%]+@[\w.\-]+\.\w+)/i, 1) || "";

    const productName = safeMatch(body, /Мероприятие:\s*([^\r\n]+)/i, 1) || "";
    const tourDateRaw = safeMatch(body, /Дата мероприятия:\s*(\d{2}\.\d{2}\.\d{4})/i, 1) || "";
    const tourDate = tourDateRaw ? formatTourDate(tourDateRaw) : "";
    const orderDate = safeMatch(body, /Дата оформления\s*(\d{2}\.\d{2}\.\d{4})/i, 1) || "";

    const qtyMatch = safeMatch(body, /Количество билетов\s*[\r\n]+\s*(\d+)/i, 1) || safeMatch(body, /Количество билетов\s*(\d+)/i, 1) || "1";
    const qtyVal = parseInt(qtyMatch, 10) || 1;
    const quantity = qtyVal.toString();

    const priceRaw = safeMatch(body, /(\d{1,3}(?:[\s\u00A0]\d{3})*|\d+)(?:[,\.]\d{2})?\s*₽/, 1);
    let totalAmount = "0";
    if (priceRaw) {
      const priceSingle = parseInt(priceRaw.replace(/\s|\u00A0/g, ""), 10);
      totalAmount = (priceSingle * qtyVal).toString();
    }

    const ticketType = "Стандарт";

    if (orderId && productName) {
      results.push({
        orderId, orderDate, firstName, lastName, phone, email,
        productName, tourDate, quantity, totalAmount, note: "", isBooking: "", ticketType
      });
    }
  } catch (e) {
    console.error("Ошибка парсинга Горбилет: " + e.message);
  }
  return results;
}

function formatTourDate(dateStr) {

  if (!dateStr) return "";

  const [day, month] = dateStr.split(".");

  const months = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];

  return `${parseInt(day, 10)}-${months[parseInt(month, 10) - 1]}`;

}


function logError(logSheet, error, messageId) {

  const now = new Date();

  logSheet.appendRow([now, messageId, error.toString()]);

}
