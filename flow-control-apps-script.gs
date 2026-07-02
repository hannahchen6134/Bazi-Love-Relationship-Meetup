const SPREADSHEET_ID = "14DPL8FAbQduZajSM-292azyTbGu2JA35xUDk86q_nJs";
const SHEET_NAME = "流程控制";

function doGet() {
  return jsonOutput_(readState_());
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const nextStep = clampStep_(payload.maxUnlockedStep);
    const updatedBy = String(payload.updatedBy || "流程控制後台");
    const updatedFrom = String(payload.updatedFrom || "web");
    const updatedAt = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");

    const sheet = getControlSheet_();
    const values = sheet.getRange("A2:B5").getValues();
    const keyToRow = new Map(values.map((row, index) => [String(row[0] || ""), index + 2]));

    sheet.getRange(`B${keyToRow.get("maxUnlockedStep") || 2}`).setValue(nextStep);
    sheet.getRange(`B${keyToRow.get("lastUpdatedAt") || 3}`).setValue(updatedAt);
    sheet.getRange(`B${keyToRow.get("lastUpdatedBy") || 4}`).setValue(updatedBy);
    sheet.getRange(`B${keyToRow.get("updatedFrom") || 5}`).setValue(updatedFrom);

    return jsonOutput_({
      ok: true,
      maxUnlockedStep: nextStep,
      updatedAt,
      updatedBy,
      updatedFrom
    });
  } catch (error) {
    return jsonOutput_({
      ok: false,
      message: error && error.message ? error.message : "Unknown error"
    });
  }
}

function readState_() {
  const sheet = getControlSheet_();
  const values = sheet.getRange("A2:B5").getValues();
  const state = Object.fromEntries(values.map((row) => [String(row[0] || ""), row[1]]));
  return {
    ok: true,
    maxUnlockedStep: clampStep_(state.maxUnlockedStep || 16),
    updatedAt: String(state.lastUpdatedAt || ""),
    updatedBy: String(state.lastUpdatedBy || ""),
    updatedFrom: String(state.updatedFrom || "")
  };
}

function getControlSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`找不到工作表：${SHEET_NAME}`);
  }
  return sheet;
}

function parsePayload_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "";
  if (!raw) {
    throw new Error("POST 內容是空的");
  }
  return JSON.parse(raw);
}

function clampStep_(value) {
  const step = Number(value);
  if (!Number.isFinite(step)) return 16;
  return Math.max(1, Math.min(16, Math.round(step)));
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
