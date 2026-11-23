/** 附近揪團｜第 1 階段 GAS API：Google 試算表資料服務 */
const EVENT_SHEET = "Events";
const JOIN_SHEET = "Joins";
const EVENT_HEADERS = ["eventId", "title", "category", "date", "time", "place", "address", "capacity", "joined", "host", "createdAt", "lat", "lng"];
const JOIN_HEADERS = ["eventId", "name", "joinedAt"];

function setupSheets() {
  const events = getSheet_(EVENT_SHEET, EVENT_HEADERS);
  const joins = getSheet_(JOIN_SHEET, JOIN_HEADERS);
  // 識別碼和文字欄位先設成純文字，避免日期、時間字串被試算表自動轉型。
  setTextColumns_(events, [[1, 7], [10, 1]]); // A:G、J
  setTextColumns_(joins, [[1, 2]]); // A:B
}

function setTextColumns_(sheet, columns) {
  const rowCount = Math.max(sheet.getMaxRows(), 2);
  columns.forEach(([startColumn, width]) => {
    sheet.getRange(1, startColumn, rowCount, width).setNumberFormat("@");
  });
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "ping";
  if (action === "ping") return json_({status: "ok", message: "Gathering API online"});
  return json_({status: "error", message: "請用 POST 呼叫資料功能"});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!authorized_(body.token)) throw new Error("API_TOKEN 不正確");
    if (body.action === "listEvents") return json_({status: "ok", events: listEvents_()});
    if (body.action === "createEvent") return json_(createEvent_(body));
    if (body.action === "joinEvent") return json_(joinEvent_(body));
    return json_({status: "error", message: "未知的 action"});
  } catch (err) {
    return json_({status: "error", message: String(err.message || err)});
  }
}

function authorized_(token) {
  const saved = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  return Boolean(saved) && String(token || "") === saved;
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function listEvents_() {
  const sheet = getSheet_(EVENT_SHEET, EVENT_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return rows.slice(1).filter(row => row[0]).map(row => ({
    id: String(row[0]), title: displayText_(row[1]), category: displayText_(row[2]),
    date: dateString_(row[3]), time: timeString_(row[4]), place: displayText_(row[5]),
    address: displayText_(row[6]), capacity: Number(row[7]), joined: Number(row[8]),
    host: displayText_(row[9]), lat: optionalNumber_(row[11]), lng: optionalNumber_(row[12])
  })).filter(item => !item.date || item.date >= today);
}

function createEvent_(body) {
  const title = clean_(body.title, 80), host = clean_(body.host, 30);
  const place = clean_(body.place, 100), address = clean_(body.address, 200);
  const capacity = Number(body.capacity);
  if (!title || !host || !place || !address) throw new Error("請填活動名稱、主揪、地標與地址");
  if (!isFinite(capacity) || capacity < 2 || capacity > 30) throw new Error("名額需介於 2 到 30 人");
  const lat = optionalNumber_(body.lat), lng = optionalNumber_(body.lng);
  if ((lat === null) !== (lng === null)) throw new Error("緯度和經度要一起填，或一起留白");
  if (lat !== null && (lat < -90 || lat > 90 || lng < -180 || lng > 180)) throw new Error("座標超出範圍");
  const date = String(body.date || "").trim(), time = String(body.time || "").trim();
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("日期格式需為 YYYY-MM-DD");
  if (time && !/^\d{2}:\d{2}$/.test(time)) throw new Error("時間格式需為 HH:MM");

  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const events = getSheet_(EVENT_SHEET, EVENT_HEADERS);
    const joins = getSheet_(JOIN_SHEET, JOIN_HEADERS);
    const id = Utilities.getUuid().slice(0, 8).toUpperCase();
    events.appendRow([
      id, safeText_(title), safeText_(clean_(body.category, 20)), safeText_(date), safeText_(time),
      safeText_(place), safeText_(address), capacity, 1, safeText_(host), new Date(),
      lat === null ? "" : lat, lng === null ? "" : lng
    ]);
    // joined 初始為 1，Joins 也記下主揪，避免主揪用同一暱稱重複報名。
    joins.appendRow([id, safeText_(host), new Date()]);
    return {status: "ok", eventId: id};
  } finally { lock.releaseLock(); }
}

function joinEvent_(body) {
  const eventId = clean_(body.eventId, 20), name = clean_(body.name, 30);
  if (!eventId || !name) throw new Error("請提供活動編號和暱稱");
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const events = getSheet_(EVENT_SHEET, EVENT_HEADERS), rows = events.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== eventId) continue;
      const capacity = Number(rows[i][7]), joined = Number(rows[i][8]);
      if (joined >= capacity) throw new Error("活動已額滿");
      const joins = getSheet_(JOIN_SHEET, JOIN_HEADERS), joinedRows = joins.getDataRange().getValues();
      if (joinedRows.slice(1).some(row => String(row[0]) === eventId && displayText_(row[1]).toLowerCase() === name.toLowerCase())) {
        throw new Error("這個暱稱已報名（主揪也已占用一個名額）");
      }
      joins.appendRow([eventId, safeText_(name), new Date()]);
      events.getRange(i + 1, 9).setValue(joined + 1);
      return {status: "ok", eventId: eventId};
    }
    throw new Error("找不到活動");
  } finally { lock.releaseLock(); }
}

function optionalNumber_(value) {
  if (value === "" || value === null || value === undefined) return null;
  const number = Number(value);
  return isFinite(number) ? number : null;
}
function clean_(value, maxLength) { return String(value || "").replace(/[<>]/g, "").trim().slice(0, maxLength); }
function safeText_(value) {
  const text = String(value || "").trim();
  // 前置單引號把 =、+、-、@ 開頭的使用者文字當成文字儲存，避免被解讀成公式。
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
function displayText_(value) {
  const text = String(value || "");
  return /^'[=+\-@]/.test(text) ? text.slice(1) : text;
}
function dateString_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(value || "").slice(0, 10);
}
function timeString_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  return String(value || "").slice(0, 5);
}
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
