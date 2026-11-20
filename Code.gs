/**
 * SATSET TECH - Badminton Tournament Manager Backend (Code.gs v10.3)
 * Kompatibel dengan index.html (Admin), public-viewer.html (Public Viewer + Registrasi + Hasil), & promotion.html
 * Update v10.1: Penyimpanan dedicated sheet _BRANDING untuk mencegah error limit 50.000 karakter sel Google Sheets
 * saat upload Favicon (icon.png), OG Banner (og-banner.png), Logo, & Watermark.
 */

const CONFIG_SHEET = '_CONFIG';
const LIVE_SHEET = '_LIVE_COURTS';
const REG_SHEET = '_PENDAFTARAN';
const RESULTS_SHEET = '_HASIL_PERTANDINGAN';
const BRACKET_SHEET = '_BAGAN_JSON';
const SPONSORS_SHEET = '_SPONSORS';
const BRANDING_SHEET = '_BRANDING';

const DEFAULT_USER = 'admin';
const DEFAULT_PASS = 'admin2026';
const DEFAULT_PIN  = '123456';

function initAuthProps() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_USER')) props.setProperty('ADMIN_USER', DEFAULT_USER);
  if (!props.getProperty('ADMIN_PASS')) props.setProperty('ADMIN_PASS', DEFAULT_PASS);
  if (!props.getProperty('MASTER_PIN')) props.setProperty('MASTER_PIN', DEFAULT_PIN);
}

function doGet(e) {
  initAuthProps();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = getConfigMeta(ss);
  const sheets = ss.getSheets();
  const sheetClasses = [];
  const participants = {};
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith('_') && name.toLowerCase() !== 'sheet1') {
      sheetClasses.push(name);
      const lastRow = sheet.getLastRow();
      const list = [];
      if (lastRow > 1) {
        const numCols = Math.max(sheet.getLastColumn(), 11);
        const rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
        rows.forEach(r => {
          if (r[0] || r[3]) {
            list.push({
              id: String(r[0]),
              player1: String(r[1]),
              player2: String(r[2]),
              name: String(r[3]),
              club: String(r[4]),
              pairType: String(r[5] || 'MD'),
              seed: r[6] !== '' && r[6] !== null ? Number(r[6]) : null,
              checkedIn: String(r[7] || '').toUpperCase() === 'HADIR',
              checkInTime: String(r[8] || ''),
              whatsapp: String(r[9] || ''),
              photoData: String(r[10] || '')
            });
          }
        });
      }
      participants[name] = list;
    }
  });

  let classes = [];
  if (meta.activeClasses && meta.activeClasses.length > 0) {
    classes = meta.activeClasses;
    classes.forEach(cls => {
      if (!participants[cls]) participants[cls] = [];
    });
  } else {
    classes = sheetClasses;
  }

  const courts = [];
  const liveSheet = ss.getSheetByName(LIVE_SHEET);
  if (liveSheet && liveSheet.getLastRow() > 1) {
    const numCols = Math.max(liveSheet.getLastColumn(), 11);
    const cRows = liveSheet.getRange(2, 1, liveSheet.getLastRow() - 1, numCols).getValues();
    cRows.forEach(r => {
      let pendingObj = null;
      if (r[7]) {
        try { pendingObj = JSON.parse(String(r[7])); } catch (err) {}
      }
      courts.push({
        court: r[0],
        status: r[1],
        className: r[2],
        matchNr: r[3],
        team1: r[4],
        team2: r[5],
        score: r[6],
        venue: String(r[8] || ''),
        oopDate: String(r[9] || ''),
        oopTime: String(r[10] || ''),
        pending: pendingObj
      });
    });
  }

  const registrations = [];
  let regSheet = ss.getSheetByName(REG_SHEET);
  if (!regSheet) {
    regSheet = ss.insertSheet(REG_SHEET);
    regSheet.appendRow(['Reg ID', 'Waktu Daftar', 'Kelas', 'Pemain 1', 'Pemain 2', 'Klub / PB', 'Tipe', 'No WhatsApp', 'Status', 'Foto Peserta']);
    regSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#ccff00');
  }
  if (regSheet && regSheet.getLastRow() > 1) {
    const numRegCols = Math.max(regSheet.getLastColumn(), 10);
    const rRows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, numRegCols).getValues();
    rRows.forEach(r => {
      if (r[0]) {
        registrations.push({
          regId: String(r[0]),
          timestamp: String(r[1]),
          className: String(r[2]),
          player1: String(r[3]),
          player2: String(r[4]),
          club: String(r[5]),
          pairType: String(r[6] || 'MD'),
          whatsapp: String(r[7] || ''),
          status: String(r[8] || 'PENDING'),
          photoData: String(r[9] || '')
        });
      }
    });
  }

  const results = [];
  const resSheet = ss.getSheetByName(RESULTS_SHEET);
  if (resSheet && resSheet.getLastRow() > 1) {
    const numResCols = Math.max(resSheet.getLastColumn(), 14);
    const resRows = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, numResCols).getValues();
    resRows.forEach(r => {
      if (r[0]) {
        results.push({
          resultId: String(r[0]),
          finishedAt: String(r[1]),
          className: String(r[2]),
          matchNr: String(r[3]),
          roundLabel: String(r[4]),
          team1: String(r[5]),
          team2: String(r[6]),
          score: String(r[7]),
          winner: String(r[8]),
          winNote: String(r[9] || 'NORMAL'),
          officials: String(r[10] || ''),
          startedTime: String(r[11] || ''),
          finishedTime: String(r[12] || ''),
          durationLabel: String(r[13] || '')
        });
      }
    });
  }

  let brackets = {};
  const bracketSheet = ss.getSheetByName(BRACKET_SHEET);
  if (bracketSheet && bracketSheet.getLastRow() > 1) {
    const bRows = bracketSheet.getRange(2, 1, bracketSheet.getLastRow() - 1, 2).getValues();
    bRows.forEach(r => {
      if (r[0] && r[1]) {
        try { brackets[String(r[0])] = JSON.parse(String(r[1])); } catch (err) {}
      }
    });
  }

  return jsonResponse({
    status: 'ok',
    eventName: meta.eventName,
    venueLocation: (meta.venueList && meta.venueList[0]) || meta.venueLocation,
    venueList: meta.venueList,
    courtVenues: meta.courtVenues,
    city: meta.city,
    dateRange: meta.dateRange,
    regOpen: meta.regOpen,
    maintenanceMode: Boolean(meta.maintenanceMode),
    paymentInfo: meta.paymentInfo,
    adminContacts: meta.adminContacts,
    themeConfig: meta.themeConfig,
    brandingConfig: meta.brandingConfig,
    sponsorsList: meta.sponsorsList,
    publicTabsConfig: meta.publicTabsConfig,
    announcementText: meta.announcementText,
    waPromoTemplate: meta.waPromoTemplate,
    handbookUrl: meta.handbookUrl,
    agencyConfig: meta.agencyConfig,
    regFieldsConfig: meta.regFieldsConfig,
    sponsorDisplayMode: meta.sponsorDisplayMode || 'AUTO',
    sponsorSpeed: Number(meta.sponsorSpeed) || 24,
    scoringRule: meta.scoringRule || '21_3SET',
    minRestMinutes: Number(meta.minRestMinutes) || 45,
    refereeName: meta.refereeName || '',
    pdfFooterNote: meta.pdfFooterNote || '',
    classes: classes,
    participants: participants,
    courts: courts,
    registrations: registrations,
    results: results,
    brackets: brackets
  });
}

function doPost(e) {
  initAuthProps();
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();

    if (action === 'verifyAuth') {
      const savedUser = props.getProperty('ADMIN_USER');
      const savedPass = props.getProperty('ADMIN_PASS');
      const savedPin  = props.getProperty('MASTER_PIN');
      const isValid = (payload.user === savedUser) && (payload.pass === savedPass || payload.pin === savedPin);
      return jsonResponse({ status: isValid ? 'ok' : 'error', message: isValid ? 'Login berhasil' : 'Kredensial salah' });
    }

    if (action === 'updateCredentials') {
      const savedPass = props.getProperty('ADMIN_PASS');
      const savedPin  = props.getProperty('MASTER_PIN');
      if (payload.oldPassOrPin !== savedPass && payload.oldPassOrPin !== savedPin) {
        return jsonResponse({ status: 'error', message: 'Verifikasi lama gagal' });
      }
      if (payload.newUser) props.setProperty('ADMIN_USER', payload.newUser);
      if (payload.newPass) props.setProperty('ADMIN_PASS', payload.newPass);
      if (payload.newMasterPin) props.setProperty('MASTER_PIN', payload.newMasterPin);
      return jsonResponse({ status: 'ok' });
    }

    if (action === 'syncClassParticipants') {
      const className = String(payload.className || 'UMUM').trim();
      const list = payload.participants || [];
      let sheet = ss.getSheetByName(className);
      if (!sheet) { sheet = ss.insertSheet(className); }
      sheet.clear();
      sheet.appendRow(['ID Pasangan', 'Pemain 1', 'Pemain 2', 'Nama Pasangan', 'Klub / PB', 'Tipe', 'Seed', 'Kehadiran', 'Jam Check-In', 'No WhatsApp', 'Foto Peserta']);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#ccff00');
      
      if (list.length > 0) {
        const rows = list.map(p => [
          p.id || '', p.player1 || '', p.player2 || '', p.name || '', p.club || '', p.pairType || 'MD', p.seed || '',
          p.checkedIn ? 'HADIR' : 'BELUM', p.checkInTime || '', p.whatsapp || '', p.photoData || ''
        ]);
        sheet.getRange(2, 1, rows.length, 11).setValues(rows);
      }

      const defaultSheet = ss.getSheetByName('Sheet1');
      if (defaultSheet && ss.getSheets().length > 1) {
        try { ss.deleteSheet(defaultSheet); } catch (err) {}
      }

      return jsonResponse({ status: 'ok', className: className, count: list.length });
    }

    if (action === 'syncLiveState') {
      setConfigMeta(ss, {
        eventName: payload.eventName || 'SATSET BADMINTON OPEN 2026',
        venueLocation: (Array.isArray(payload.venueList) && payload.venueList[0]) || payload.venueLocation || 'GOR PLATINUM ARENA MALANG',
        venueList: Array.isArray(payload.venueList) && payload.venueList.length > 0 ? payload.venueList : [payload.venueLocation || 'GOR PLATINUM ARENA MALANG'],
        courtVenues: payload.courtVenues || {},
        city: payload.city || 'MALANG, INA',
        dateRange: payload.dateRange || '08-13 Nov 2026',
        regOpen: payload.regOpen || 'OPEN',
        maintenanceMode: payload.maintenanceMode !== undefined ? Boolean(payload.maintenanceMode) : undefined,
        paymentInfo: payload.paymentInfo !== undefined ? payload.paymentInfo : undefined,
        adminContacts: Array.isArray(payload.adminContacts) ? payload.adminContacts : undefined,
        themeConfig: payload.themeConfig !== undefined ? payload.themeConfig : undefined,
        brandingConfig: payload.brandingConfig !== undefined ? payload.brandingConfig : undefined,
        sponsorsList: Array.isArray(payload.sponsorsList) ? payload.sponsorsList : undefined,
        publicTabsConfig: payload.publicTabsConfig !== undefined ? payload.publicTabsConfig : undefined,
        announcementText: payload.announcementText !== undefined ? payload.announcementText : undefined,
        waPromoTemplate: payload.waPromoTemplate !== undefined ? payload.waPromoTemplate : undefined,
        handbookUrl: payload.handbookUrl !== undefined ? payload.handbookUrl : undefined,
        agencyConfig: payload.agencyConfig !== undefined ? payload.agencyConfig : undefined,
        regFieldsConfig: payload.regFieldsConfig !== undefined ? payload.regFieldsConfig : undefined,
        sponsorDisplayMode: payload.sponsorDisplayMode !== undefined ? payload.sponsorDisplayMode : undefined,
        sponsorSpeed: payload.sponsorSpeed !== undefined ? Number(payload.sponsorSpeed) : undefined,
        scoringRule: payload.scoringRule !== undefined ? payload.scoringRule : undefined,
        minRestMinutes: payload.minRestMinutes !== undefined ? Number(payload.minRestMinutes) : undefined,
        refereeName: payload.refereeName !== undefined ? payload.refereeName : undefined,
        pdfFooterNote: payload.pdfFooterNote !== undefined ? payload.pdfFooterNote : undefined,
        activeClasses: Array.isArray(payload.activeClasses) ? payload.activeClasses : null
      });

      let liveSheet = ss.getSheetByName(LIVE_SHEET);
      if (!liveSheet) liveSheet = ss.insertSheet(LIVE_SHEET);
      liveSheet.clear();
      liveSheet.appendRow(['Court', 'Status', 'Kelas', 'No Partai', 'Team 1', 'Team 2', 'Skor Live', 'Pending Next Match', 'Lokasi Lapangan', 'Tanggal OOP', 'Jam OOP']);
      liveSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#0b0f17').setFontColor('#ccff00');

      const courts = payload.courts || [];
      if (courts.length > 0) {
        const cRows = courts.map(c => [
          c.court, c.status || 'STANDBY', c.className || '-', c.matchNr || '-', c.team1 || '-', c.team2 || '-', c.score || '-',
          c.pending ? JSON.stringify(c.pending) : '', c.venue || '', c.oopDate || '', c.oopTime || ''
        ]);
        liveSheet.getRange(2, 1, cRows.length, 11).setValues(cRows);
      }

      if (Array.isArray(payload.results)) {
        let resSheet = ss.getSheetByName(RESULTS_SHEET);
        if (!resSheet) resSheet = ss.insertSheet(RESULTS_SHEET);
        resSheet.clear();
        resSheet.appendRow(['Result ID', 'Waktu Selesai', 'Kelas', 'No Partai', 'Babak', 'Team 1', 'Team 2', 'Skor', 'Pemenang', 'Keterangan', 'Perangkat Pertandingan', 'Mulai', 'Selesai', 'Durasi']);
        resSheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#0b0f17').setFontColor('#ccff00');
        if (payload.results.length > 0) {
          const rRows = payload.results.map(r => [
            r.resultId || '', r.finishedAt || '', r.className || '', r.matchNr || '', r.roundLabel || '',
            r.team1 || '', r.team2 || '', r.score || '', r.winner || '', r.winNote || 'NORMAL', r.officials || '',
            r.startedTime || '', r.finishedTime || '', r.durationLabel || ''
          ]);
          resSheet.getRange(2, 1, rRows.length, 14).setValues(rRows);
        }
      }

      if (payload.brackets && typeof payload.brackets === 'object') {
        let bSheet = ss.getSheetByName(BRACKET_SHEET);
        if (!bSheet) bSheet = ss.insertSheet(BRACKET_SHEET);
        bSheet.clear();
        bSheet.appendRow(['Class', 'BracketJSON']);
        const bEntries = Object.keys(payload.brackets).map(k => [k, JSON.stringify(payload.brackets[k])]);
        if (bEntries.length > 0) {
          bSheet.getRange(2, 1, bEntries.length, 2).setValues(bEntries);
        }
      }

      return jsonResponse({ status: 'ok' });
    }

    if (action === 'submitRegistration') {
      const meta = getConfigMeta(ss);
      if (meta.regOpen === 'CLOSED') {
        return jsonResponse({ status: 'error', message: 'Maaf, pendaftaran online sudah ditutup oleh panitia.' });
      }
      let regSheet = ss.getSheetByName(REG_SHEET);
      if (!regSheet) {
        regSheet = ss.insertSheet(REG_SHEET);
        regSheet.appendRow(['Reg ID', 'Waktu Daftar', 'Kelas', 'Pemain 1', 'Pemain 2', 'Klub / PB', 'Tipe', 'No WhatsApp', 'Status', 'Foto Peserta']);
        regSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#ccff00');
      }
      const regId = 'REG-' + new Date().getTime().toString().slice(-7);
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      regSheet.appendRow([
        regId,
        nowStr,
        String(payload.className || 'UMUM').trim(),
        String(payload.player1 || '').trim(),
        String(payload.player2 || '').trim(),
        String(payload.club || 'UMUM').trim().toUpperCase(),
        String(payload.pairType || 'MD').trim(),
        String(payload.whatsapp || '').trim(),
        'PENDING',
        String(payload.photoData || '')
      ]);
      return jsonResponse({ status: 'ok', regId: regId });
    }

    if (action === 'resetAllTournamentData') {
      const sheets = ss.getSheets();
      sheets.forEach(sh => {
        const sName = sh.getName();
        if (sName === LIVE_SHEET || sName === RESULTS_SHEET || sName === BRACKET_SHEET || sName === REG_SHEET) {
          if (sh.getLastRow() > 1) {
            sh.deleteRows(2, sh.getLastRow() - 1);
          }
        } else if (!sName.startsWith('_')) {
          if (ss.getSheets().length > 1) {
            try { ss.deleteSheet(sh); } catch (err) {}
          } else {
            sh.clear();
          }
        }
      });
      return jsonResponse({ status: 'ok', message: 'Semua data turnamen berhasil direset bersih.' });
    }

    if (action === 'archiveClassSheet') {
      const className = String(payload.className || '').trim();
      if (className) {
        const sheet = ss.getSheetByName(className);
        if (sheet) {
          const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmm');
          const archiveName = '_ARSIP_' + className + '_' + stamp;
          sheet.setName(archiveName);
        }
      }
      return jsonResponse({ status: 'ok', archived: className });
    }

    if (action === 'updateRegistrationStatus') {
      const regSheet = ss.getSheetByName(REG_SHEET);
      if (regSheet && regSheet.getLastRow() > 1) {
        const lastRow = regSheet.getLastRow();
        const ids = regSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(payload.regId)) {
            regSheet.getRange(i + 2, 9).setValue(payload.status || 'APPROVED');
            break;
          }
        }
      }
      return jsonResponse({ status: 'ok' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function getConfigMeta(ss) {
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  const meta = {
    eventName: 'SATSET BADMINTON OPEN 2026',
    venueLocation: 'GOR PLATINUM ARENA MALANG',
    venueList: ['GOR PLATINUM ARENA MALANG'],
    city: 'MALANG, INA',
    dateRange: 'Min, 08 - Jum, 13 Nov 2026',
    regOpen: 'OPEN',
    maintenanceMode: false,
    paymentInfo: '',
    adminContacts: [],
    activeClasses: [],
    courtVenues: {},
    themeConfig: null,
    brandingConfig: null,
    sponsorsList: [],
    publicTabsConfig: null,
    announcementText: '',
    waPromoTemplate: '',
    handbookUrl: '',
    agencyConfig: null,
    regFieldsConfig: null,
    sponsorDisplayMode: 'AUTO',
    sponsorSpeed: 24,
    scoringRule: '21_3SET',
    minRestMinutes: 45,
    refereeName: '',
    pdfFooterNote: ''
  };
  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    rows.forEach(r => {
      if (r[0] === 'eventName') meta.eventName = r[1];
      if (r[0] === 'venueLocation') meta.venueLocation = r[1];
      if (r[0] === 'venueList' && r[1]) {
        try { meta.venueList = JSON.parse(r[1]); } catch (e) {}
      }
      if (r[0] === 'city') meta.city = r[1];
      if (r[0] === 'dateRange') meta.dateRange = r[1];
      if (r[0] === 'regOpen') meta.regOpen = r[1];
      if (r[0] === 'maintenanceMode') meta.maintenanceMode = (r[1] === true || String(r[1]).toUpperCase() === 'TRUE');
      if (r[0] === 'paymentInfo') meta.paymentInfo = r[1];
      if (r[0] === 'adminContacts' && r[1]) { try { meta.adminContacts = JSON.parse(r[1]); } catch (e) {} }
      if (r[0] === 'themeConfig' && r[1]) { try { meta.themeConfig = JSON.parse(r[1]); } catch (e) {} }
      if (r[0] === 'brandingConfig' && r[1] && r[1] !== ('STORED_IN_' + BRANDING_SHEET)) {
        try { meta.brandingConfig = JSON.parse(r[1]); } catch (e) {}
      }
      if (r[0] === 'sponsorsList' && r[1] && r[1] !== ('STORED_IN_' + SPONSORS_SHEET)) {
        try { meta.sponsorsList = JSON.parse(r[1]); } catch (e) {}
      }
      if (r[0] === 'publicTabsConfig' && r[1]) { try { meta.publicTabsConfig = JSON.parse(r[1]); } catch (e) {} }
      if (r[0] === 'announcementText') meta.announcementText = String(r[1] || '');
      if (r[0] === 'waPromoTemplate') meta.waPromoTemplate = String(r[1] || '');
      if (r[0] === 'handbookUrl') meta.handbookUrl = String(r[1] || '');
      if (r[0] === 'agencyConfig' && r[1]) { try { meta.agencyConfig = JSON.parse(r[1]); } catch (e) {} }
      if (r[0] === 'regFieldsConfig' && r[1]) { try { meta.regFieldsConfig = JSON.parse(r[1]); } catch (e) {} }
      if (r[0] === 'sponsorDisplayMode' && r[1]) meta.sponsorDisplayMode = String(r[1]);
      if (r[0] === 'sponsorSpeed' && r[1] !== '') meta.sponsorSpeed = Number(r[1]) || 24;
      if (r[0] === 'scoringRule' && r[1]) meta.scoringRule = String(r[1]);
      if (r[0] === 'minRestMinutes' && r[1] !== '') meta.minRestMinutes = Number(r[1]) || 45;
      if (r[0] === 'refereeName') meta.refereeName = String(r[1] || '');
      if (r[0] === 'pdfFooterNote') meta.pdfFooterNote = String(r[1] || '');
      if (r[0] === 'courtVenues' && r[1]) {
        try { meta.courtVenues = JSON.parse(r[1]); } catch (e) {}
      }
      if (r[0] === 'activeClasses' && r[1]) {
        try { meta.activeClasses = JSON.parse(r[1]); } catch (e) {}
      }
    });
  }

  // Load dedicated _BRANDING sheet (supports large base64 icon.png, og-banner.png, logoData, watermarkData per row)
  const brSheet = ss.getSheetByName(BRANDING_SHEET);
  if (brSheet && brSheet.getLastRow() > 1) {
    const brRows = brSheet.getRange(2, 1, brSheet.getLastRow() - 1, 2).getValues();
    const loadedBranding = meta.brandingConfig && typeof meta.brandingConfig === 'object' ? Object.assign({}, meta.brandingConfig) : {};
    brRows.forEach(br => {
      const k = String(br[0] || '');
      const v = br[1];
      if (!k) return;
      if (typeof v === 'string' && (v.startsWith('{') || v.startsWith('[') || v === 'true' || v === 'false' || v === 'null')) {
        try { loadedBranding[k] = JSON.parse(v); } catch (e) { loadedBranding[k] = v; }
      } else {
        loadedBranding[k] = v;
      }
    });
    meta.brandingConfig = loadedBranding;
  }

  const spSheet = ss.getSheetByName(SPONSORS_SHEET);
  if (spSheet && spSheet.getLastRow() > 1) {
    const spRows = spSheet.getRange(2, 1, spSheet.getLastRow() - 1, 4).getValues();
    const loadedSponsors = [];
    spRows.forEach(sr => {
      if (sr[3]) {
        loadedSponsors.push({
          name: String(sr[1] || ''),
          tier: String(sr[2] || ''),
          logoData: String(sr[3] || '')
        });
      }
    });
    meta.sponsorsList = loadedSponsors;
  }
  return meta;
}

function setConfigMeta(ss, meta) {
  const existing = getConfigMeta(ss);
  const classesToSave = (meta.activeClasses && meta.activeClasses.length > 0) ? meta.activeClasses : existing.activeClasses;
  const venuesToSave = (meta.courtVenues && typeof meta.courtVenues === 'object') ? meta.courtVenues : existing.courtVenues;
  const venueListToSave = (Array.isArray(meta.venueList) && meta.venueList.length > 0) ? meta.venueList : existing.venueList;
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET);
  sheet.clear();
  sheet.appendRow(['Key', 'Value']);
  sheet.appendRow(['eventName', meta.eventName]);
  sheet.appendRow(['venueLocation', (venueListToSave && venueListToSave[0]) || meta.venueLocation || 'GOR PLATINUM ARENA MALANG']);
  sheet.appendRow(['venueList', JSON.stringify(venueListToSave || ['GOR PLATINUM ARENA MALANG'])]);
  sheet.appendRow(['courtVenues', JSON.stringify(venuesToSave || {})]);
  sheet.appendRow(['city', meta.city]);
  sheet.appendRow(['dateRange', meta.dateRange]);
  sheet.appendRow(['regOpen', meta.regOpen || 'OPEN']);
  sheet.appendRow(['maintenanceMode', meta.maintenanceMode !== undefined ? Boolean(meta.maintenanceMode) : Boolean(existing.maintenanceMode)]);
  sheet.appendRow(['paymentInfo', meta.paymentInfo || '']);
  sheet.appendRow(['adminContacts', JSON.stringify(meta.adminContacts || existing.adminContacts || [])]);
  sheet.appendRow(['themeConfig', JSON.stringify(meta.themeConfig !== undefined ? meta.themeConfig : (existing.themeConfig || null))]);

  // Save brandingConfig per-key into dedicated _BRANDING sheet so each asset (favicon, OG banner, logo, watermark) gets its own 50k-char cell budget
  const brandingToSave = meta.brandingConfig !== undefined ? meta.brandingConfig : (existing.brandingConfig || null);
  let brSheet = ss.getSheetByName(BRANDING_SHEET);
  if (!brSheet) brSheet = ss.insertSheet(BRANDING_SHEET);
  brSheet.clear();
  brSheet.appendRow(['BrandingKey', 'BrandingValue']);
  brSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#0b0f17').setFontColor('#ccff00');
  if (brandingToSave && typeof brandingToSave === 'object') {
    const brKeys = Object.keys(brandingToSave);
    if (brKeys.length > 0) {
      const brRows = brKeys.map(k => {
        const val = brandingToSave[k];
        const serialized = (val !== null && typeof val === 'object') ? JSON.stringify(val) : (val === undefined ? '' : val);
        return [k, typeof serialized === 'string' && serialized.length > 49500 ? serialized.slice(0, 49500) : serialized];
      });
      brSheet.getRange(2, 1, brRows.length, 2).setValues(brRows);
    }
  }
  sheet.appendRow(['brandingConfig', 'STORED_IN_' + BRANDING_SHEET]);

  const sponsorsToSave = meta.sponsorsList !== undefined ? meta.sponsorsList : (existing.sponsorsList || []);
  let spSheet = ss.getSheetByName(SPONSORS_SHEET);
  if (!spSheet) spSheet = ss.insertSheet(SPONSORS_SHEET);
  spSheet.clear();
  spSheet.appendRow(['No', 'Nama Sponsor', 'Kategori / Tier', 'Logo Data']);
  spSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0b0f17').setFontColor('#ccff00');
  if (Array.isArray(sponsorsToSave) && sponsorsToSave.length > 0) {
    const spRows = sponsorsToSave.map((sp, i) => [
      i + 1,
      String(sp.name || ''),
      String(sp.tier || ''),
      String(sp.logoData || '')
    ]);
    spSheet.getRange(2, 1, spRows.length, 4).setValues(spRows);
  }
  sheet.appendRow(['sponsorsList', 'STORED_IN_' + SPONSORS_SHEET]);
  sheet.appendRow(['publicTabsConfig', JSON.stringify(meta.publicTabsConfig !== undefined ? meta.publicTabsConfig : (existing.publicTabsConfig || null))]);
  sheet.appendRow(['announcementText', meta.announcementText !== undefined ? meta.announcementText : (existing.announcementText || '')]);
  sheet.appendRow(['waPromoTemplate', meta.waPromoTemplate !== undefined ? meta.waPromoTemplate : (existing.waPromoTemplate || '')]);
  sheet.appendRow(['handbookUrl', meta.handbookUrl !== undefined ? meta.handbookUrl : (existing.handbookUrl || '')]);
  sheet.appendRow(['agencyConfig', JSON.stringify(meta.agencyConfig !== undefined ? meta.agencyConfig : (existing.agencyConfig || null))]);
  sheet.appendRow(['regFieldsConfig', JSON.stringify(meta.regFieldsConfig !== undefined ? meta.regFieldsConfig : (existing.regFieldsConfig || null))]);
  sheet.appendRow(['sponsorDisplayMode', meta.sponsorDisplayMode !== undefined ? meta.sponsorDisplayMode : (existing.sponsorDisplayMode || 'AUTO')]);
  sheet.appendRow(['sponsorSpeed', meta.sponsorSpeed !== undefined ? Number(meta.sponsorSpeed) : (Number(existing.sponsorSpeed) || 24)]);
  sheet.appendRow(['scoringRule', meta.scoringRule !== undefined ? meta.scoringRule : (existing.scoringRule || '21_3SET')]);
  sheet.appendRow(['minRestMinutes', meta.minRestMinutes !== undefined ? Number(meta.minRestMinutes) : (Number(existing.minRestMinutes) || 45)]);
  sheet.appendRow(['refereeName', meta.refereeName !== undefined ? meta.refereeName : (existing.refereeName || '')]);
  sheet.appendRow(['pdfFooterNote', meta.pdfFooterNote !== undefined ? meta.pdfFooterNote : (existing.pdfFooterNote || '')]);
  sheet.appendRow(['activeClasses', JSON.stringify(classesToSave || [])]);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
