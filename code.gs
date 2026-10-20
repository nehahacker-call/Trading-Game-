const SPREADSHEET_ID = "1Y10pcb7iSPMXDf14xdBIbUmDweKH9gtacQy9pl090q0";
const SESSION_SECONDS = 21600;
const USERS_SHEET = "Users";
const PRODUCTS_SHEET = "Products";
const HOLDINGS_SHEET = "Holdings";
const TRANSACTIONS_SHEET = "Transactions";

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Games · Share Market")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function validateUserId(userId) {
  const wantedId = normalizeId_(userId);
  if (!wantedId) return { valid: false };

  const user = findUser_(wantedId);
  if (!user || !user.active) return { valid: false };

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("session:" + token, wantedId, SESSION_SECONDS);
  return { valid: true, token: token, userId: user.displayId };
}

function testUserIdLogin() {
  const testUserId = "5555";
  const login = validateUserId(testUserId);
  if (!login.valid) {
    Logger.log("LOGIN FAILED: User ID not found or inactive.");
    return false;
  }

  const dashboard = getDashboard(login.token);
  Logger.log("LOGIN OK: balance=" + dashboard.balance + ", products=" + dashboard.products.length);
  return true;
}

function getDashboard(token) {
  const user = requireUser_(token);
  const ss = db_();
  const products = readProducts_(ss);
  const holdingSheet = requireSheet_(ss, HOLDINGS_SHEET);
  const holdingRows = holdingSheet.getDataRange().getValues();
  const holdings = holdingRows.slice(1)
    .filter(row => normalizeId_(row[0]) === user.userId && Number(row[2]) > 0)
    .map(row => ({
      productId: String(row[1] || ""),
      quantity: Number(row[2]) || 0,
      averageBuyPrice: Number(row[3]) || 0
    }));

  return {
    userId: user.displayId,
    balance: user.balance,
    products: products,
    holdings: holdings
  };
}

function buyShare(token, productId, quantity) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const user = requireUser_(token);
    const product = findProduct_(productId);
    const qty = validQuantity_(quantity);
    const total = product.buyPrice * qty;
    if (product.buyPrice <= 0) throw new Error("এই product-ৰ buy price সঠিক নহয়");
    if (user.balance < total) throw new Error("Balance যথেষ্ট নহয়");

    const nextBalance = user.balance - total;
    const ss = db_();
    const holdingSheet = requireSheet_(ss, HOLDINGS_SHEET);
    const holding = findHolding_(holdingSheet, user.userId, product.productId);
    if (holding) {
      const oldQuantity = Number(holding.values[2]) || 0;
      const oldAverage = Number(holding.values[3]) || 0;
      const nextQuantity = oldQuantity + qty;
      const nextAverage = ((oldQuantity * oldAverage) + total) / nextQuantity;
      holdingSheet.getRange(holding.row, 3, 1, 2).setValues([[nextQuantity, nextAverage]]);
    } else {
      holdingSheet.appendRow([user.userId, product.productId, qty, product.buyPrice]);
    }

    writeBalance_(user, nextBalance);
    addTransaction_(user, product, "buy", qty, product.buyPrice, total, nextBalance);
    return { ok: true, balance: nextBalance };
  } finally {
    lock.releaseLock();
  }
}

function sellShare(token, productId, quantity) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const user = requireUser_(token);
    const product = findProduct_(productId);
    const qty = validQuantity_(quantity);
    if (product.sellPrice <= 0) throw new Error("এই product-ৰ sell price সঠিক নহয়");

    const ss = db_();
    const holdingSheet = requireSheet_(ss, HOLDINGS_SHEET);
    const holding = findHolding_(holdingSheet, user.userId, product.productId);
    if (!holding || Number(holding.values[2]) < qty) {
      throw new Error("ইমান share তোমাৰ account-ত নাই");
    }

    const nextQuantity = Number(holding.values[2]) - qty;
    holdingSheet.getRange(holding.row, 3).setValue(nextQuantity);
    const total = product.sellPrice * qty;
    const nextBalance = user.balance + total;
    writeBalance_(user, nextBalance);
    addTransaction_(user, product, "sell", qty, product.sellPrice, total, nextBalance);
    return { ok: true, balance: nextBalance };
  } finally {
    lock.releaseLock();
  }
}

function db_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_YOUR_GOOGLE_SHEET_ID_HERE") {
    throw new Error("Code.gs-ত Google Sheet ID বহুৱাব লাগিব");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function requireUser_(token) {
  const userId = CacheService.getScriptCache().get("session:" + String(token || ""));
  if (!userId) throw new Error("Login session শেষ হৈছে; পুনৰ login কৰা");

  const user = findUser_(userId);
  if (!user || !user.active) throw new Error("Active user পোৱা নাই");
  return user;
}

function findUser_(wantedId) {
  const sheet = requireSheet_(db_(), USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const headers = (values[0] || []).map(value => String(value || "").trim().toLowerCase());
  const displayValues = sheet.getDataRange().getDisplayValues();
  const userIdColumn = findColumn_(headers, ["userid", "user id"], 0);
  const balanceColumn = findColumn_(headers, ["balance"], headers.length >= 3 ? 2 : 1);
  const statusColumn = findColumn_(headers, ["status"], -1);

  for (let i = 1; i < values.length; i++) {
    const displayId = String(displayValues[i][userIdColumn] || "").trim();
    if (normalizeId_(displayId) !== wantedId) continue;
    const status = statusColumn < 0 ? "active" : String(values[i][statusColumn] || "active").trim().toLowerCase();
    const rawBalance = values[i][balanceColumn];
    const balance = rawBalance === "" || rawBalance == null ? 0 : Number(rawBalance);
    if (!Number.isFinite(balance) || balance < 0) throw new Error("Users tab-ত এই user-ৰ balance সঠিক নহয়");
    return {
      row: i + 1,
      userId: wantedId,
      displayId: displayId,
      balanceColumn: balanceColumn + 1,
      balance: balance,
      active: status === "" || status === "active"
    };
  }
  return null;
}

function writeBalance_(user, balance) {
  requireSheet_(db_(), USERS_SHEET).getRange(user.row, user.balanceColumn).setValue(balance);
}

function readProducts_(ss) {
  return requireSheet_(ss, PRODUCTS_SHEET).getDataRange().getValues().slice(1)
    .filter(row => row[0])
    .map(row => ({
      productId: String(row[0]).trim(),
      name: String(row[1] || ""),
      style: String(row[2] || ""),
      buyPrice: Number(row[3]) || 0,
      sellPrice: Number(row[4]) || 0,
      imageURL: String(row[5] || "")
    }));
}

function findProduct_(productId) {
  const product = readProducts_(db_()).find(item =>
    item.productId.toLowerCase() === String(productId || "").trim().toLowerCase()
  );
  if (!product) throw new Error("Product পোৱা নাই");
  return product;
}

function findHolding_(sheet, userId, productId) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalizeId_(rows[i][0]) === userId &&
        String(rows[i][1] || "").trim().toLowerCase() === productId.toLowerCase()) {
      return { row: i + 1, values: rows[i] };
    }
  }
  return null;
}

function addTransaction_(user, product, action, quantity, unitPrice, total, balanceAfter) {
  requireSheet_(db_(), TRANSACTIONS_SHEET).appendRow([
    Utilities.getUuid(),
    user.displayId,
    product.productId,
    product.name,
    action,
    quantity,
    unitPrice,
    total,
    balanceAfter,
    new Date()
  ]);
}

function requireSheet_(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(name + " tab পোৱা নাই");
  return sheet;
}

function findColumn_(headers, candidates, fallback) {
  const index = headers.findIndex(header => candidates.includes(header.replace(/[^a-z]/g, "")));
  return index < 0 ? fallback : index;
}

function normalizeId_(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function validQuantity_(quantity) {
  const value = Number(quantity);
  if (!Number.isInteger(value) || value < 1) {
    throw new Error("Quantity 1 বা তাতকৈ বেছি পূৰ্ণ সংখ্যা হ’ব লাগিব");
  }
  return value;
}
