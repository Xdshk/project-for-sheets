function processOrderEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Автомат") || ss.insertSheet("Автомат");
  const logSheet = ss.getSheetByName("Errors") || ss.insertSheet("Errors");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const existingKeys = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastCol)
        .getValues()
        .map(r => normalizeKey(`${r[0] || ""}_${r[11] || ""}_${r[5] || ""}_${r[2] || ""}_${r[16] || ""}`))
    : [];

  const threads = GmailApp.search('subject:(Московиты) newer_than:1d');
  const now = new Date();
  const DAY = 24 * 60 * 60 * 1000;

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      try {
        if (now - message.getDate() > DAY) continue;
        const body = message.getPlainBody();
        const orders = parseOrderEmail(body);

        console.log("Письмо " + message.getId() + ": найдено билетов - " + orders.length);

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
            "www",              // 15
            order.ticketType,   // 16
            order.note || "",   // 17
            order.isBooking || "" // 18
          ]);

          existingKeys.push(uniqueKey);
        }
      } catch (e) {
        logError(logSheet, e, message.getId());
      }
    }
  }
}

function parseOrderEmail(body) {
  const results = [];
  if (!body || typeof body !== "string") return results;

  try {
    const orderId = safeMatch(body, /Заказ №\s*(\d+)/i, 1) || "";
    const orderDate = safeMatch(body, /\((\d{2}\.\d{2}\.\d{4})\)/, 1) || "";
    const nameFull = safeMatch(body, /Поступил заказ от\s+([^\n:]+)/i, 1) || "";
    const nameParts = nameFull.trim() ? nameFull.trim().split(/\s+/) : [];
    const firstName = nameParts.length ? nameParts[0] : "";
    const lastName = nameParts.length > 1 ? nameParts.slice(1).join(" ") : "";

    let phoneRaw = safeMatch(body, /\b(\+7|7|8)\d{10}\b/, 0) || "";
    let phone = normalizePhone(phoneRaw);

    const email = safeMatch(body, /\b[\w.-]+@[\w.-]+\.\w+\b/, 0) || "";
    const note = body.includes("Онлайн оплата") ? "" : "";
    const isBooking = body.includes("Предварительная запись") ? "бронирование" : "";

    const ticketBlocks = body.split(/Тип билета\s*:/i);

    for (let i = 1; i < ticketBlocks.length; i++) {
      const currentBlock = ticketBlocks[i];
      const prevBlock = ticketBlocks[i - 1];

      const ticketType = (currentBlock.split("\n")[0] || "").trim() || "Стандарт";

      const prevLines = prevBlock.split("\n")
        .map(l => l.trim())
        .filter(l => l && !l.startsWith("http") && !l.includes("wc-orders"));

      const lastFewLines = prevLines.slice(-4).join(" ");
      let productName = cleanProductName(lastFewLines);

      const dateMatchRaw = safeMatch(currentBlock, /Дата и время:\s*(\d{2}\.\d{2}\.\d{4})/i, 1);
      const tourDate = dateMatchRaw ? formatTourDate(dateMatchRaw) : "";

      const qtyMatch = safeMatch(currentBlock, /×\s*(\d+)/, 1) || "1";
      const quantity = qtyMatch || "1";

      const priceRaw = safeMatch(currentBlock, /(\d{1,3}(?:[\s\u00A0]\d{3})*|\d+)(?:[,\.]\d{2})?\s*₽/, 1);
      let totalAmount = "0";
      if (priceRaw) {
        const priceNumber = parseInt(priceRaw.replace(/\s|\u00A0/g, ""), 10);
        totalAmount = (priceNumber * parseInt(quantity, 10)).toString();
      }

      results.push({
        orderId, orderDate, firstName, lastName, phone, email,
        productName, tourDate, quantity, totalAmount, note, isBooking, ticketType
      });
    }
  } catch (e) {
    console.error("Ошибка парсинга: " + e.message);
  }
  return results;
}


function formatReserveDate(dateStr) {

  if (!dateStr) return "";



  const monthsMap = {

    "янв": "янв", "фев": "фев", "мар": "мар", "апр": "апр", "май": "май", "июн": "июн",

    "июл": "июл", "авг": "авг", "сен": "сен", "окт": "окт", "ноя": "ноя", "дек": "дек",

    "января": "янв", "февраля": "фев", "марта": "мар", "апреля": "апр", "мая": "май", "июня": "июн",

    "июля": "июл", "августа": "авг", "сентября": "сен", "октября": "окт", "ноября": "ноя", "декабря": "дек"

  };

  const dayMatch = dateStr.match(/\d+/);

  if (!dayMatch) return dateStr;

  const day = parseInt(dayMatch[0], 10);

  let foundMonth = "";

  const lowerDate = dateStr.toLowerCase();

  for (let key in monthsMap) {

    if (lowerDate.includes(key)) {

      foundMonth = monthsMap[key];

      break;

    }

  }



  return foundMonth ? `${day}-${foundMonth}` : dateStr;

}

function logError(logSheet, error, messageId) {

  const now = new Date();

  logSheet.appendRow([now, messageId, error.toString()]);

}
