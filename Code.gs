/**
 * SATSET TECH — Google Apps Script Backend (v6.4)
 * Mendukung Sinkronisasi Kartu 1 - 14 (termasuk Mode Sponsor, TTS, Aturan Skor 42 Poin, Jeda Istirahat & Tanda Tangan PDF)
 */

const SHEET_STATE = "LIVE_STATE";
const SHEET_REGISTRATIONS = "REGISTRATIONS";

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "GET_LIVE_STATE";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "GET_LIVE_STATE") {
    const sheet = getOrCreateSheet_(ss, SHEET_STATE);
    const rawJson = sheet.getRange("A2").getValue();
    if (!rawJson) {
      return jsonResponse_({ ok: true, empty: true });
    }
    return ContentService.createTextOutput(rawJson).setMimeType(ContentService.MimeType.JSON);
  }

  return jsonResponse_({ ok: true, message: "SATSET TECH API Active" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === "SYNC_LIVE_STATE") {
      const sheet = getOrCreateSheet_(ss, SHEET_STATE);
      const statePayload = {
        tournamentMeta: body.tournamentMeta || {},
        sponsorsList: body.sponsorsList || [],
        sponsorDisplayMode: body.sponsorDisplayMode || "AUTO",
        sponsorSpeed: body.sponsorSpeed || 24,
        scoringRule: body.scoringRule || "21_3SET",
        minRestMinutes: body.minRestMinutes || 45,
        refereeName: body.refereeName || "",
        pdfFooterNote: body.pdfFooterNote || "",
        matchesList: body.matchesList || [],
        updatedAt: new Date().toISOString()
      };

      sheet.getRange("A1:D1").setValues([["STATE_JSON", "SCORING_RULE", "SPONSOR_MODE", "UPDATED_AT"]]);
      sheet.getRange("A2:D2").setValues([[
        JSON.stringify(statePayload),
        statePayload.scoringRule,
        statePayload.sponsorDisplayMode,
        statePayload.updatedAt
      ]]);

      return jsonResponse_({ ok: true, updatedAt: statePayload.updatedAt });
    }

    if (body.action === "REGISTER_TEAM") {
      const regSheet = getOrCreateSheet_(ss, SHEET_REGISTRATIONS);
      if (regSheet.getLastRow() === 0) {
        regSheet.appendRow(["TIMESTAMP", "KELAS", "PEMAIN_1", "PEMAIN_2", "KLUB", "WHATSAPP", "STATUS"]);
      }
      regSheet.appendRow([
        new Date(),
        body.classCode || "",
        body.player1 || "",
        body.player2 || "",
        body.club || "",
        body.whatsapp || "",
        "PENDING"
      ]);
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.toString() });
  }
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
