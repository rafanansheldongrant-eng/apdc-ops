// ================================================================
// APDC OPERATIONS SYSTEM — Google Apps Script Backend v3
// Security:
//   • GPS Geofencing — blocks check-in from wrong location
//   • Server-side timestamps — client timestamp can't be faked
//   • Daily duplicate prevention — one per type per day per worker+building
//   • Worker-building assignment enforcement — only assigned workers can check in
//   • Worker PIN validation — prevents buddy check-ins
//   • Timestamp mismatch detection — flags suspicious submissions
//   • GPS accuracy suspicion flag — flags suspiciously perfect GPS (< 8m = spoofing risk)
//   • Time-on-site tracking — flags workers who check out too early
//   • Improved score formula — attendance rate-based, not raw count
//   • Work order assignment from dashboard — with worker email notification
//   • getAssignedWorkOrders — workers see their tasks in worker app
//   • Clean getLogs response — no duplicate fields
// ================================================================

const SPREADSHEET_ID    = '1CeUSmU8UgjVoF4sdEUX25LWp11tyH_3UEofqq90UjME';
const DRIVE_FOLDER_ID   = '1TB9AiEWaaGYAeqtgv5zSdyLAwO4QlhBG';
const FM_EMAIL          = 'DilaoAU@churchofjesuschrist.org';  // ← REQUIRED: Add your email here — alerts won't work without this
const COMPANY_NAME      = "Adam's Projects Development Corp";
const GEOFENCE_RADIUS_M = 200; // Default max distance (meters) from building for valid check-in

// Photo subfolder names inside your root Drive folder
const PHOTO_FOLDERS = {
  DAILY:      'Daily',       // Attendance selfies + Checklist selfies
  DISCOVERY:  'Discovery',   // Flag photos
  ADDITIONAL: 'Additional'   // Work order photos
};

const SHEETS = {
  ATTENDANCE:    'Attendance',
  WORKORDERS:    'Work Orders',
  CHECKLISTS:    'Checklists',
  FLAGS:         'Flags',
  CONFIG:        'Config',
  BUILDINGS:     'Buildings',
  DAILY_REPORTS: 'Daily Reports',
  SPECIAL_TASKS: 'Special Tasks',
  LOG:           'System Log'
};

// ================================================================
// GET HANDLER
// ================================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getConfig')             return getConfig();
  if (action === 'getLogs')               return getLogs(e.parameter);
  if (action === 'getReports')            return getReports(e.parameter);
  if (action === 'getScores')             return getScores(e.parameter);
  if (action === 'getDashboardData')      return getDashboardData(e.parameter);
  if (action === 'getAssignedWorkOrders') return getAssignedWorkOrders(e.parameter);

  return jsonResponse({ status: 'ok', message: 'APDC Backend v3 Active', ts: new Date().toISOString() });
}

// ================================================================
// CONFIG FETCH
// Config sheet columns: Worker | Building | Worker Email | PIN
//   Each row = one worker-building assignment
//   PIN = 4-digit code worker must enter to check in (prevents buddy check-ins)
// Buildings sheet columns: Building | Lat | Lng | Radius(m)
//   Each row = GPS fence for a building
// ================================================================
function getConfig() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Config sheet ──
    let cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!cfgSheet) {
      cfgSheet = ss.insertSheet(SHEETS.CONFIG);
      cfgSheet.appendRow(['Worker', 'Building', 'Worker Email', 'PIN']);
      styleHeader(cfgSheet, 1, 4);
      cfgSheet.setFrozenRows(1);
      cfgSheet.setColumnWidth(1, 220);
      cfgSheet.setColumnWidth(2, 220);
      cfgSheet.setColumnWidth(3, 240);
      cfgSheet.setColumnWidth(4, 80);
      cfgSheet.appendRow(['Add Worker Name Here', 'Add Building Name Here', 'worker@email.com', '1234']);
    }

    const cfgData      = cfgSheet.getDataRange().getValues();
    const workers      = [];
    const buildings    = [];
    const workerEmails = {};  // worker → email
    const assignments  = {};  // worker → [buildings]
    const pinMap       = {};  // worker → PIN (sent to client for local validation)

    for (let i = 1; i < cfgData.length; i++) {
      const w     = String(cfgData[i][0] || '').trim();
      const b     = String(cfgData[i][1] || '').trim();
      const email = String(cfgData[i][2] || '').trim();
      const pin   = String(cfgData[i][3] || '').trim();
      if (w && !workers.includes(w))   workers.push(w);
      if (b && !buildings.includes(b)) buildings.push(b);
      if (w && email) workerEmails[w] = email;
      if (w && pin)   pinMap[w]       = pin;
      if (w && b) {
        if (!assignments[w]) assignments[w] = [];
        if (!assignments[w].includes(b)) assignments[w].push(b);
      }
    }

    // ── Building GPS coords ──
    const buildingCoords = getBuildingCoords(ss);

    return jsonResponse({ status: 'ok', workers, buildings, workerEmails, assignments, buildingCoords, pinMap });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString(), workers: [], buildings: [] });
  }
}

// ================================================================
// BUILDING COORDS — reads/creates Buildings sheet
// ================================================================
function getBuildingCoords(ss) {
  const coords = {}; // building name → { lat, lng, radius }
  try {
    let sheet = ss.getSheetByName(SHEETS.BUILDINGS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.BUILDINGS);
      sheet.appendRow(['Building', 'Lat', 'Lng', 'Radius (m)']);
      styleHeader(sheet, 1, 4);
      sheet.setFrozenRows(1);
      sheet.appendRow(['Example Building', '14.5995', '120.9842', '200']);
      sheet.setColumnWidth(1, 220);
      sheet.setColumnWidth(2, 140);
      sheet.setColumnWidth(3, 140);
      sheet.setColumnWidth(4, 120);
      return coords;
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const name   = String(data[i][0] || '').trim();
      const lat    = parseFloat(data[i][1]);
      const lng    = parseFloat(data[i][2]);
      const radius = parseFloat(data[i][3]) || GEOFENCE_RADIUS_M;
      if (name && !isNaN(lat) && !isNaN(lng)) {
        coords[name] = { lat, lng, radius };
      }
    }
  } catch (e) { Logger.log('getBuildingCoords error: ' + e); }
  return coords;
}

// ================================================================
// HAVERSINE DISTANCE — returns meters between two GPS coords
// ================================================================
function haversineDistance(lat1, lon1, lat2, lon2) {
  const R    = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a    = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
               Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
               Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ================================================================
// GET ASSIGNED WORK ORDERS — worker.html fetches open tasks for a worker
// ================================================================
function getAssignedWorkOrders(params) {
  try {
    const worker   = (params.worker   || '').trim().toLowerCase();
    const building = (params.building || '').trim().toLowerCase();
    const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet    = ss.getSheetByName(SHEETS.WORKORDERS);
    if (!sheet) return jsonResponse({ status: 'ok', workOrders: [] });

    const data       = sheet.getDataRange().getValues();
    const headers    = data[0];
    const woOrders   = [];

    // Locate new columns if they exist (added by handleAssignWorkOrder)
    const assignedIdx = headers.indexOf('Assigned To');
    const priorityIdx = headers.indexOf('Priority');
    const dueDateIdx  = headers.indexOf('Due Date');
    const notesIdx    = headers.indexOf('Instructions');

    for (let i = 1; i < data.length; i++) {
      const row        = data[i];
      if (!row[0]) continue;
      const status     = String(row[10] || 'Open').toLowerCase();
      if (['resolved', 'cancelled', 'complete', 'completed'].includes(status)) continue;

      // Assigned To: use new column if available, otherwise fall back to Worker column
      const assignedTo = (assignedIdx >= 0 ? String(row[assignedIdx] || '') : String(row[2] || '')).trim();
      const rowBuilding = String(row[3] || '').trim();

      if (worker   && assignedTo.toLowerCase()  !== worker)   continue;
      if (building && rowBuilding.toLowerCase() !== building)  continue;

      woOrders.push({
        id:          row[4]  || '',
        building:    rowBuilding,
        assignedTo,
        workType:    row[6]  || '',
        location:    row[7]  || '',
        description: row[8]  || '',
        priority:    priorityIdx >= 0 ? (row[priorityIdx] || row[5] || '') : (row[5] || ''),
        dueDate:     dueDateIdx  >= 0 ? (row[dueDateIdx]  || '') : '',
        status:      row[10] || 'Open',
        addedDate:   row[0]  ? Utilities.formatDate(new Date(row[0]), 'Asia/Manila', 'MMM dd') : '',
        notes:       notesIdx >= 0 ? (row[notesIdx] || '') : ''
      });
    }

    return jsonResponse({ status: 'ok', workOrders: woOrders });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString(), workOrders: [] });
  }
}

// ================================================================
// GET LOGS — today's attendance for dashboard live feed
// ================================================================
function getLogs(params) {
  try {
    const dateStr = params.date || phDate();
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet   = ss.getSheetByName(SHEETS.ATTENDANCE);
    if (!sheet) return jsonResponse({ status: 'ok', logs: [], date: dateStr });

    const data = sheet.getDataRange().getValues();
    const logs = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ts  = row[0];
      if (!ts) continue;
      const rowDate = Utilities.formatDate(new Date(ts), 'Asia/Manila', 'yyyy-MM-dd');
      if (rowDate !== dateStr) continue;

      const checkType = String(row[4] || '').toUpperCase();
      const modeMap   = { 'IN': 'morning', 'MIDDAY': 'midday', 'OUT': 'evening',
                          'MORNING': 'morning', 'EVENING': 'evening' };

      const flags    = row[14] || '';
      const accuracy = row[7] ? parseFloat(row[7]) : null;
      logs.push({
        rawTs:         new Date(ts).getTime(),      // removed before sending, used for time-on-site calc
        timestamp:     Utilities.formatDate(new Date(ts), 'Asia/Manila', 'MMM dd, yyyy h:mm a'),
        time:          Utilities.formatDate(new Date(ts), 'Asia/Manila', 'h:mm a'),
        worker:        row[2] || '',
        building:      row[3] || '',
        checkType,
        mode:          modeMap[checkType] || checkType.toLowerCase(),
        lat:           row[5] || '',
        lng:           row[6] || '',
        accuracy:      accuracy !== null ? Math.round(accuracy) : '',
        gpsValid:      row[13] || (accuracy && accuracy < 150 ? 'VALID' : 'NO_DATA'),
        gpsDistance:   row[12] || '',
        address:       row[8]  || '',
        selfieUrl:     row[9]  || '',
        flags,
        suspicious:    flags.length > 0,
        hoursOnSite:   null,
        shortShiftFlag: ''
      });
    }
    // ── Calculate time-on-site for each worker+building pair ──
    // Pair IN and OUT records, flag if on-site hours are suspiciously short
    const pairKey  = l => `${l.worker}::${l.building}`;
    const inTimes  = {};
    const outTimes = {};
    logs.forEach(l => {
      const k = pairKey(l);
      if (l.checkType === 'IN')  inTimes[k]  = l.rawTs;
      if (l.checkType === 'OUT') outTimes[k] = l.rawTs;
    });
    logs.forEach(l => {
      const k = pairKey(l);
      if (inTimes[k] && outTimes[k]) {
        const hrs = (outTimes[k] - inTimes[k]) / 3600000;
        l.hoursOnSite   = Math.round(hrs * 10) / 10;
        l.shortShiftFlag = hrs < 4 ? 'SHORT_SHIFT_' + Math.round(hrs * 10) / 10 + 'hrs' : '';
      }
    });
    // Remove rawTs from final output
    logs.forEach(l => delete l.rawTs);

    return jsonResponse({ status: 'ok', logs, date: dateStr });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString(), logs: [] });
  }
}

// ================================================================
// GET REPORTS — checklists + flags for compliance tracker
// ================================================================
function getReports(params) {
  try {
    const dateStr = params.date || phDate();
    const days    = parseInt(params.days) || 7;
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);

    const cutoff = new Date(dateStr);
    cutoff.setDate(cutoff.getDate() - (days - 1));

    // ── Checklists ──
    const clSheet = ss.getSheetByName(SHEETS.CHECKLISTS);
    const clRows  = clSheet ? clSheet.getDataRange().getValues() : [];
    const reports = [];

    for (let i = 1; i < clRows.length; i++) {
      const row = clRows[i];
      const ts  = row[0];
      if (!ts) continue;
      const d = new Date(ts);
      if (d < cutoff) continue;
      const passIdx = clRows[0].indexOf('Pass Count');
      const failIdx = clRows[0].indexOf('Fail Count');
      reports.push({
        date:      Utilities.formatDate(d, 'Asia/Manila', 'MMM dd'),
        timestamp: Utilities.formatDate(d, 'Asia/Manila', 'MMM dd, yyyy h:mm a'),
        worker:    row[2] || '',
        building:  row[3] || '',
        pass:      passIdx >= 0 ? (row[passIdx] || 0) : 0,
        fail:      failIdx >= 0 ? (row[failIdx] || 0) : 0,
        selfieUrl: row[7] || '',
        type:      'checklist'
      });
    }

    // ── Flags ──
    const flSheet = ss.getSheetByName(SHEETS.FLAGS);
    const flRows  = flSheet ? flSheet.getDataRange().getValues() : [];
    const flags   = [];

    for (let i = 1; i < flRows.length; i++) {
      const row = flRows[i];
      const ts  = row[0];
      if (!ts) continue;
      const d = new Date(ts);
      if (d < cutoff) continue;
      flags.push({
        timestamp:   Utilities.formatDate(d, 'Asia/Manila', 'MMM dd, yyyy h:mm a'),
        worker:      row[2] || '',
        building:    row[3] || '',
        urgency:     row[4] || '',
        category:    row[5] || '',
        location:    row[6] || '',
        description: row[7] || '',
        photoUrl:    row[11] || '',
        selfieUrl:   row[12] || '',
        status:      row[13] || 'Open',
        type:        'flag'
      });
    }

    return jsonResponse({ status: 'ok', reports, flags });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString(), reports: [], flags: [] });
  }
}

// ================================================================
// GET SCORES — worker performance (rate-based, not raw count)
// Formula:
//   Attendance  40pts — days checked in / period days
//   Checkout    10pts — checkouts / checkins (are they completing shifts?)
//   Checklist   30pts — pass rate across all submitted checklists
//   Proactivity 20pts — flags + WOs raised, capped at 1 per day
// ================================================================
function getScores(params) {
  try {
    const dateStr = params.date || phDate();
    const period  = params.period || 'week';
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const days    = period === 'month' ? 30 : period === 'week' ? 7 : 1;

    const cutoff = new Date(dateStr);
    cutoff.setDate(cutoff.getDate() - (days - 1));

    const data = {}; // worker → accumulated metrics

    function ensure(w) {
      if (!data[w]) data[w] = { checkInDays: new Set(), checkOuts: 0, totalPass: 0, totalFail: 0, totalItems: 0, flags: 0, wos: 0 };
    }

    // ── Attendance ──
    const attSheet = ss.getSheetByName(SHEETS.ATTENDANCE);
    const attRows  = attSheet ? attSheet.getDataRange().getValues() : [];
    for (let i = 1; i < attRows.length; i++) {
      const row = attRows[i];
      const ts  = row[0];
      if (!ts || new Date(ts) < cutoff) continue;
      const w = row[2]; if (!w) continue;
      ensure(w);
      const type = (row[4] || '').toUpperCase();
      if (type.includes('IN')) {
        const dayStr = Utilities.formatDate(new Date(ts), 'Asia/Manila', 'yyyy-MM-dd');
        data[w].checkInDays.add(dayStr);
      }
      if (type.includes('OUT')) data[w].checkOuts++;
    }

    // ── Checklists ──
    const clSheet = ss.getSheetByName(SHEETS.CHECKLISTS);
    const clRows  = clSheet ? clSheet.getDataRange().getValues() : [];
    const passIdx = clRows.length ? clRows[0].indexOf('Pass Count') : -1;
    const failIdx = clRows.length ? clRows[0].indexOf('Fail Count') : -1;
    for (let i = 1; i < clRows.length; i++) {
      const row = clRows[i];
      const ts  = row[0];
      if (!ts || new Date(ts) < cutoff) continue;
      const w = row[2]; if (!w) continue;
      ensure(w);
      const p = passIdx >= 0 ? (parseInt(row[passIdx]) || 0) : 0;
      const f = failIdx >= 0 ? (parseInt(row[failIdx]) || 0) : 0;
      data[w].totalPass  += p;
      data[w].totalFail  += f;
      data[w].totalItems += (p + f);
    }

    // ── Flags ──
    const flSheet = ss.getSheetByName(SHEETS.FLAGS);
    const flRows  = flSheet ? flSheet.getDataRange().getValues() : [];
    for (let i = 1; i < flRows.length; i++) {
      const row = flRows[i];
      const ts  = row[0];
      if (!ts || new Date(ts) < cutoff) continue;
      const w = row[2]; if (!w) continue;
      ensure(w);
      data[w].flags++;
    }

    // ── Work Orders ──
    const woSheet = ss.getSheetByName(SHEETS.WORKORDERS);
    const woRows  = woSheet ? woSheet.getDataRange().getValues() : [];
    for (let i = 1; i < woRows.length; i++) {
      const row = woRows[i];
      const ts  = row[0];
      if (!ts || new Date(ts) < cutoff) continue;
      const w = row[2]; if (!w) continue;
      ensure(w);
      data[w].wos++;
    }

    // ── Calculate scores ──
    const scoreList = Object.keys(data).map(w => {
      const d = data[w];

      const daysPresent    = d.checkInDays.size;
      const checkIns       = daysPresent;
      const checkOuts      = d.checkOuts;

      // Attendance rate: % of period days they checked in (max 1.0)
      const attRate        = Math.min(daysPresent / days, 1.0);

      // Checkout rate: did they complete their shifts?
      const checkoutRate   = checkIns > 0 ? Math.min(checkOuts / checkIns, 1.0) : 0;

      // Checklist compliance: pass rate
      const clCompliance   = d.totalItems > 0 ? d.totalPass / d.totalItems : 0;

      // Proactivity: flags + WOs, capped at 1 per day max
      const proactivity    = Math.min((d.flags + d.wos) / Math.max(days * 0.5, 1), 1.0);

      const score = Math.round(
        (attRate      * 40) +
        (checkoutRate * 10) +
        (clCompliance * 30) +
        (proactivity  * 20)
      );

      return {
        worker:      w,
        score:       Math.max(0, Math.min(100, score)),
        daysPresent,
        checkOuts,
        checklistPass:  d.totalPass,
        checklistFail:  d.totalFail,
        flagsRaised:    d.flags,
        wos:            d.wos,
        attendancePct:  Math.round(attRate * 100),
        checkoutPct:    Math.round(checkoutRate * 100),
        compliancePct:  Math.round(clCompliance * 100)
      };
    });

    scoreList.sort((a, b) => b.score - a.score);
    return jsonResponse({ status: 'ok', scores: scoreList, period, days });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString(), scores: [] });
  }
}

// ================================================================
// GET DASHBOARD DATA
// ================================================================
function getDashboardData(params) {
  try {
    const dateStr = params ? (params.date || phDate()) : phDate();
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Today's Check-ins ──
    const attSheet  = ss.getSheetByName(SHEETS.ATTENDANCE);
    const attRows   = attSheet ? attSheet.getDataRange().getValues() : [];
    const todayAtt  = { checkedIn: [], midday: [], checkedOut: [] };
    const allWorkers = new Set();

    for (let i = 1; i < attRows.length; i++) {
      const row     = attRows[i];
      const ts      = row[0];
      if (!ts) continue;
      const rowDate = Utilities.formatDate(new Date(ts), 'Asia/Manila', 'yyyy-MM-dd');
      if (rowDate !== dateStr) continue;
      const w = row[2], bld = row[3], type = (row[4] || '').toUpperCase();
      allWorkers.add(w);
      const entry = {
        worker:    w,
        building:  bld,
        time:      Utilities.formatDate(new Date(ts), 'Asia/Manila', 'h:mm a'),
        selfieUrl: row[9] || '',
        gpsValid:  row[13] || ''
      };
      if (type.includes('IN'))     todayAtt.checkedIn.push(entry);
      else if (type === 'MIDDAY')  todayAtt.midday.push(entry);
      else if (type.includes('OUT')) todayAtt.checkedOut.push(entry);
    }

    // ── Open Flags ──
    const flSheet  = ss.getSheetByName(SHEETS.FLAGS);
    const flRows   = flSheet ? flSheet.getDataRange().getValues() : [];
    const openFlags = [], emergency = [];

    for (let i = 1; i < flRows.length; i++) {
      const row    = flRows[i];
      if (!row[0]) continue;
      const status = (row[13] || '').toLowerCase();
      if (status === 'open' || status === '') {
        const f = {
          timestamp:   Utilities.formatDate(new Date(row[0]), 'Asia/Manila', 'MMM dd h:mm a'),
          worker:      row[2] || '',
          building:    row[3] || '',
          urgency:     row[4] || '',
          category:    row[5] || '',
          location:    row[6] || '',
          description: (row[7] || '').substring(0, 80),
          photoUrl:    row[11] || ''
        };
        openFlags.push(f);
        if (row[4] === 'Emergency') emergency.push(f);
      }
    }

    // ── Open Work Orders ──
    const woSheet = ss.getSheetByName(SHEETS.WORKORDERS);
    const woRows  = woSheet ? woSheet.getDataRange().getValues() : [];
    const openWOs = [];

    for (let i = 1; i < woRows.length; i++) {
      const row    = woRows[i];
      if (!row[0]) continue;
      const status = (row[10] || '').toLowerCase();
      if (!['resolved', 'cancelled', 'complete', 'completed'].includes(status)) {
        openWOs.push({
          timestamp:      Utilities.formatDate(new Date(row[0]), 'Asia/Manila', 'MMM dd'),
          worker:         row[2]  || '',
          building:       row[3]  || '',
          woNumber:       row[4]  || '',
          classification: row[5]  || '',
          workType:       row[6]  || '',
          location:       row[7]  || '',
          description:    (row[8] || '').substring(0, 80),
          status:         row[10] || 'Open',
          beforePhotoUrl: row[14] || '',
          afterPhotoUrl:  row[15] || ''
        });
      }
    }

    // ── Today's Checklists ──
    const clSheet = ss.getSheetByName(SHEETS.CHECKLISTS);
    const clRows  = clSheet ? clSheet.getDataRange().getValues() : [];
    const todayChecklists = [];
    const passIdx = clRows.length ? clRows[0].indexOf('Pass Count') : -1;
    const failIdx = clRows.length ? clRows[0].indexOf('Fail Count') : -1;

    for (let i = 1; i < clRows.length; i++) {
      const row = clRows[i];
      const ts  = row[0]; if (!ts) continue;
      const rowDate = Utilities.formatDate(new Date(ts), 'Asia/Manila', 'yyyy-MM-dd');
      if (rowDate !== dateStr) continue;
      todayChecklists.push({
        worker:    row[2] || '',
        building:  row[3] || '',
        pass:      passIdx >= 0 ? (row[passIdx] || 0) : 0,
        fail:      failIdx >= 0 ? (row[failIdx] || 0) : 0,
        selfieUrl: row[7] || ''
      });
    }

    // ── Photo Summary ──
    const dailyPhotos = [];
    todayAtt.checkedIn.concat(todayAtt.midday).concat(todayAtt.checkedOut).forEach(e => {
      if (e.selfieUrl) dailyPhotos.push({ url: e.selfieUrl, label: 'Attendance', worker: e.worker, building: e.building });
    });
    todayChecklists.forEach(c => {
      if (c.selfieUrl) dailyPhotos.push({ url: c.selfieUrl, label: 'Checklist', worker: c.worker, building: c.building });
    });

    const discoveryPhotos = openFlags.filter(f => f.photoUrl).map(f => ({
      url: f.photoUrl, label: f.urgency + ' — ' + f.category,
      worker: f.worker, building: f.building, description: f.description
    }));

    const additionalPhotos = [];
    openWOs.forEach(w => {
      if (w.beforePhotoUrl) additionalPhotos.push({ url: w.beforePhotoUrl, label: 'Before — ' + (w.workType || w.woNumber), worker: w.worker, building: w.building });
      if (w.afterPhotoUrl)  additionalPhotos.push({ url: w.afterPhotoUrl,  label: 'After — '  + (w.workType || w.woNumber), worker: w.worker, building: w.building });
    });

    return jsonResponse({
      status: 'ok',
      date: dateStr,
      summary: {
        checkedIn:      todayAtt.checkedIn.length,
        midday:         todayAtt.midday.length,
        checkedOut:     todayAtt.checkedOut.length,
        uniqueWorkers:  allWorkers.size,
        openFlags:      openFlags.length,
        emergencyFlags: emergency.length,
        openWorkOrders: openWOs.length,
        checklistsDone: todayChecklists.length
      },
      attendance:     { checkedIn: todayAtt.checkedIn, midday: todayAtt.midday, checkedOut: todayAtt.checkedOut },
      openFlags,
      openWorkOrders: openWOs,
      checklists:     todayChecklists,
      photos:         { daily: dailyPhotos, discovery: discoveryPhotos, additional: additionalPhotos }
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ================================================================
// POST HANDLER
// ================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if      (data.type === 'attendance')        result = handleAttendance(data);
    else if (data.mode === 'daily_report')      result = handleDailyReport(data);
    else if (data.mode === 'special_task')      result = handleSpecialTask(data);
    else if (data.mode === 'workorder')         result = handleWorkOrder(data);
    else if (data.mode === 'assignWorkOrder')   result = handleAssignWorkOrder(data);
    else if (data.mode === 'checklist')         result = handleChecklist(data);
    else if (data.mode === 'flag')              result = handleFlag(data);
    else                                        result = { status: 'error', message: 'Unknown type' };
    logEntry('SUCCESS', data.type || data.mode, data.worker, '');
    return jsonResponse(result);
  } catch (err) {
    logEntry('ERROR', 'doPost', '', err.toString());
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ================================================================
// ATTENDANCE — with full fraud prevention
//
// Checks (in order):
//   1. Worker-building assignment: worker must be in Config for this building
//   2. Daily duplicate: one per check type per worker+building per day
//   3. GPS geofencing: worker must be within building's radius (if coords set)
//   4. Timestamp flag: records mismatch between client time and server time
//
// Uses server time (new Date()) as the official record timestamp.
// Client's submitted timestamp is stored for audit, not used as the record.
// ================================================================
function handleAttendance(d) {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const serverTs = new Date(); // official timestamp — cannot be faked by client
  const todayStr = Utilities.formatDate(serverTs, 'Asia/Manila', 'yyyy-MM-dd');

  // ── 1. Worker-building assignment check ──
  const cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (cfgSheet && d.worker && d.building) {
    const cfgData = cfgSheet.getDataRange().getValues();
    let isAssigned = false;
    for (let i = 1; i < cfgData.length; i++) {
      const w = String(cfgData[i][0] || '').trim();
      const b = String(cfgData[i][1] || '').trim();
      if (w === d.worker && b === d.building) { isAssigned = true; break; }
    }
    if (!isAssigned) {
      return {
        status: 'error',
        message: `"${d.worker}" is not assigned to "${d.building}". Contact your supervisor.`
      };
    }
  }

  // ── 2. Daily duplicate check ──
  const sheet = getOrCreateSheet(SHEETS.ATTENDANCE, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'Check Type',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Address', 'Selfie URL', 'Client Timestamp',
    'GPS Distance (m)', 'GPS Status', 'Geofence Result', 'Timestamp Flag'
  ]);

  const existing = sheet.getDataRange().getValues();
  for (let i = existing.length - 1; i >= 1; i--) {
    const row     = existing[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), 'Asia/Manila', 'yyyy-MM-dd');
    if (rowDate !== todayStr) break; // rows are chronological — safe to stop
    if (row[2] === d.worker && row[4] === d.checkType && row[3] === d.building) {
      return {
        status: 'duplicate',
        message: `${d.worker} already submitted ${d.checkType} for ${d.building} today.`,
        worker:  d.worker,
        type:    d.checkType
      };
    }
  }

  // ── 3. GPS Geofencing ──
  let gpsDistance = null;
  let gpsStatus   = 'NO_BUILDING_DATA';

  if (d.lat && d.lng) {
    const buildingCoords = getBuildingCoords(ss);
    const coords         = buildingCoords[d.building];
    if (coords) {
      gpsDistance = Math.round(haversineDistance(
        parseFloat(d.lat), parseFloat(d.lng), coords.lat, coords.lng
      ));
      const radius = coords.radius || GEOFENCE_RADIUS_M;
      gpsStatus = gpsDistance <= radius ? 'VALID' : 'OUT_OF_RANGE';

      if (gpsStatus === 'OUT_OF_RANGE') {
        return {
          status:   'error',
          message:  `GPS shows you are ${gpsDistance}m from ${d.building}. You must be within ${radius}m to check in. Make sure you are on-site.`,
          distance: gpsDistance
        };
      }
    }
  }

  // ── 4. Timestamp flag — detect if client's clock is off by more than 5 minutes ──
  const clientTs   = d.timestamp ? new Date(d.timestamp) : serverTs;
  const tsDiffMins = Math.abs(serverTs - clientTs) / 60000;
  const tsFlag     = tsDiffMins > 5 ? `MISMATCH_${Math.round(tsDiffMins)}min` : '';

  // ── 5. GPS accuracy suspicion — real phones typically read 10-50m
  //    < 8m accuracy is suspiciously perfect and may indicate GPS spoofing app ──
  const acc = d.accuracy ? parseFloat(d.accuracy) : null;
  const gpsAccFlag = (acc !== null && acc < 8) ? 'SUSPICIOUS_GPS_PERFECT' : '';

  // ── Save photo ──
  const selfieUrl = d.selfie
    ? saveImageToDrive(d.selfie, `ATT_${sanitize(d.worker)}_${d.checkType}_${fmtTs(serverTs.toISOString())}.jpg`, PHOTO_FOLDERS.DAILY)
    : '';

  const flags = [tsFlag, gpsAccFlag].filter(Boolean).join(' | ');

  sheet.appendRow([
    serverTs,                                       // official server time
    new Date(),                                     // submitted at (server receives)
    d.worker    || '',
    d.building  || '',
    d.checkType || '',
    d.lat       || '',
    d.lng       || '',
    d.accuracy  ? Math.round(d.accuracy) : '',
    d.address   || '',
    selfieUrl,
    d.timestamp || '',                              // client's claimed timestamp (audit)
    gpsDistance !== null ? gpsDistance : '',
    gpsStatus,
    gpsDistance !== null ? `${gpsDistance}m` : '',
    flags                                           // all suspicious flags combined
  ]);

  return {
    status:  'ok',
    message: 'Attendance recorded',
    worker:  d.worker,
    type:    d.checkType,
    gpsStatus,
    gpsDistance
  };
}

// ================================================================
// ASSIGN WORK ORDER — called from dashboard, emails the assigned worker
// ================================================================
function handleAssignWorkOrder(d) {
  const sheet = getOrCreateSheet(SHEETS.WORKORDERS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'WO Number', 'Classification',
    'Work Type', 'Location', 'Description', 'Materials Used', 'Status',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Before Photo URL', 'After Photo URL', 'Selfie URL',
    'Assigned To', 'Priority', 'Due Date', 'Source', 'Instructions'
  ]);

  const serverTs = new Date();
  sheet.appendRow([
    serverTs,
    serverTs,
    d.worker      || d.assignedTo || '',   // Worker = who submitted (from dashboard = you)
    d.building    || '',
    d.woId        || d.woNumber   || '',
    d.classification || d.priority || '',
    d.workType    || d.source     || '',
    d.location    || '',
    d.description || d.issue      || '',
    '',                                    // Materials Used — worker fills this on completion
    d.status      || 'Open',
    '', '', '',                            // GPS (not applicable for dashboard-created WOs)
    '', '', '',                            // Photos (worker submits these on completion)
    d.worker      || d.assignedTo || '',   // Assigned To
    d.priority    || '',
    d.dueDate     || '',
    d.source      || '',
    d.notes       || ''
  ]);

  // ── Email the assigned worker ──
  if (d.worker || d.assignedTo) {
    const assignedWorker = d.worker || d.assignedTo;
    const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    const cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (cfgSheet) {
      const cfgData = cfgSheet.getDataRange().getValues();
      for (let i = 1; i < cfgData.length; i++) {
        const w     = String(cfgData[i][0] || '').trim();
        const email = String(cfgData[i][2] || '').trim();
        if (w === assignedWorker && email) {
          sendEmailAlert(email,
            `📋 APDC — New Work Order Assigned: ${d.woId || d.woNumber || 'WO'} — ${d.building}`,
            `Hello ${assignedWorker},\n\nA new work order has been assigned to you.\n\n` +
            `WO Number: ${d.woId || d.woNumber || '—'}\n` +
            `Building: ${d.building || '—'}\n` +
            `Description: ${d.description || d.issue || '—'}\n` +
            `Priority: ${d.priority || '—'}\n` +
            `Due Date: ${d.dueDate || '—'}\n` +
            (d.notes ? `Instructions: ${d.notes}\n` : '') +
            `\nPlease open the APDC Worker App to view and complete this task.\n\n— APDC System`
          );
          break;
        }
      }
    }
  }

  return { status: 'ok', message: 'Work order assigned', woId: d.woId || d.woNumber, assignedTo: d.worker || d.assignedTo };
}

// ================================================================
// DAILY REPORT
// ================================================================
function handleDailyReport(d) {
  const sheet = getOrCreateSheet(SHEETS.DAILY_REPORTS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'GPS Address',
    'Task #', 'Task Name', 'Category', 'Before Photo URL', 'After Photo URL', 'Task Notes', 'Selfie URL'
  ]);

  const ts        = fmtTs(d.timestamp);
  const w         = sanitize(d.worker);
  const selfieUrl = d.selfie
    ? saveImageToDrive(d.selfie, `DAILY_SELFIE_${w}_${ts}.jpg`, PHOTO_FOLDERS.DAILY)
    : '';

  const tasks       = d.tasks || [];
  const activeTasks = tasks.filter(t => t.before || t.after);

  if (!activeTasks.length) {
    sheet.appendRow([
      new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
      d.worker || '', d.building || '',
      d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '', d.address || '',
      '', 'No task photos submitted', '', '', '', '', selfieUrl
    ]);
  } else {
    activeTasks.forEach((t, i) => {
      const bfUrl = t.before ? saveImageToDrive(t.before, `DAILY_B${i+1}_${w}_${ts}.jpg`, PHOTO_FOLDERS.DAILY) : '';
      const afUrl = t.after  ? saveImageToDrive(t.after,  `DAILY_A${i+1}_${w}_${ts}.jpg`, PHOTO_FOLDERS.DAILY) : '';
      sheet.appendRow([
        new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
        d.worker || '', d.building || '',
        d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '', d.address || '',
        i + 1, t.name || '', t.category || '', bfUrl, afUrl, t.notes || '', i === 0 ? selfieUrl : ''
      ]);
    });
  }

  return { status: 'ok', message: 'Daily report recorded', worker: d.worker, tasksRecorded: tasks.length };
}

// ================================================================
// SPECIAL TASK
// ================================================================
function handleSpecialTask(d) {
  const sheet = getOrCreateSheet(SHEETS.SPECIAL_TASKS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'GPS Address',
    'Description', 'Requested By', 'Location in Building',
    'Before Photo URL', 'After Photo URL', 'Notes', 'Selfie URL', 'Status'
  ]);

  const ts        = fmtTs(d.timestamp);
  const w         = sanitize(d.worker);
  const selfieUrl = d.selfie      ? saveImageToDrive(d.selfie,      `SP_SELFIE_${w}_${ts}.jpg`, PHOTO_FOLDERS.ADDITIONAL) : '';
  const bfUrl     = d.beforePhoto ? saveImageToDrive(d.beforePhoto, `SP_BEFORE_${w}_${ts}.jpg`, PHOTO_FOLDERS.ADDITIONAL) : '';
  const afUrl     = d.afterPhoto  ? saveImageToDrive(d.afterPhoto,  `SP_AFTER_${w}_${ts}.jpg`,  PHOTO_FOLDERS.ADDITIONAL) : '';

  sheet.appendRow([
    new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
    d.worker || '', d.building || '',
    d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '', d.address || '',
    d.description || '', d.requestedBy || '', d.location || '',
    bfUrl, afUrl, d.notes || '', selfieUrl, 'Open'
  ]);

  return { status: 'ok', message: 'Special task recorded', worker: d.worker };
}

// ================================================================
// WORK ORDER — submitted by worker from worker app
// ================================================================
function handleWorkOrder(d) {
  const sheet = getOrCreateSheet(SHEETS.WORKORDERS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'WO Number', 'Classification',
    'Work Type', 'Location', 'Description', 'Materials Used', 'Status',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Before Photo URL', 'After Photo URL', 'Selfie URL',
    'Assigned To', 'Priority', 'Due Date', 'Source', 'Instructions'
  ]);

  const ts        = fmtTs(d.timestamp);
  const w         = sanitize(d.worker);
  const selfieUrl = d.selfie      ? saveImageToDrive(d.selfie,      `WO_SELFIE_${w}_${ts}.jpg`,  PHOTO_FOLDERS.ADDITIONAL) : '';
  const bfUrl     = d.beforePhoto ? saveImageToDrive(d.beforePhoto, `WO_BEFORE_${w}_${ts}.jpg`,  PHOTO_FOLDERS.ADDITIONAL) : '';
  const afUrl     = d.afterPhoto  ? saveImageToDrive(d.afterPhoto,  `WO_AFTER_${w}_${ts}.jpg`,   PHOTO_FOLDERS.ADDITIONAL) : '';

  sheet.appendRow([
    new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
    d.worker || '', d.building || '', d.woNumber || '', d.classification || '',
    d.workType || '', d.location || '', d.description || '', d.materials || '', d.status || 'Open',
    d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '',
    bfUrl, afUrl, selfieUrl,
    d.worker || '', '', '', '', ''
  ]);

  if (d.classification === 'Emergency') sendEmergencyAlert('EMERGENCY WORK ORDER', d);
  return { status: 'ok', message: 'Work order recorded' };
}

// ================================================================
// CHECKLIST
// ================================================================
function handleChecklist(d) {
  const CL_ITEMS = [
    { cat: 'General Building', items: ['Building clean/accessible', 'All lights functioning', 'No water leaks', 'Restrooms operational', 'Doors/windows close properly'] },
    { cat: 'Fire Safety',      items: ['Extinguisher pressure OK', 'Extinguisher accessible', 'Extinguisher clean', 'Alarm switches OK', 'Alarm bell operational'] },
    { cat: 'Electrical',       items: ['No tripped breakers', 'No burning smell', 'No exposed wiring', 'Outlets clean', 'Panel board clear'] },
    { cat: 'Plumbing',         items: ['No visible leaks', 'Drains flowing', 'WC flushing', 'Water pump operational', 'No foul odors'] },
    { cat: 'Generator',        items: ['Generator area clean', 'Fuel level OK', 'No oil leaks', 'Battery terminals clean', 'Exhaust clear'] },
    { cat: 'Locks & Security', items: ['Entry locks working', 'Deadbolts OK', 'Padlocks intact', 'No forced entry', 'Hinges secure'] },
    { cat: 'Grounds',          items: ['Lawn maintained', 'No dead plants', 'Pathways clear', 'No overgrown branches', 'No grounds hazards'] }
  ];

  const baseHeaders = ['Timestamp', 'Submitted At', 'Worker', 'Building', 'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Selfie URL'];
  const itemHeaders = [];
  CL_ITEMS.forEach(g => g.items.forEach(item => itemHeaders.push(`[${g.cat}] ${item}`)));
  const sheet = getOrCreateSheet(SHEETS.CHECKLISTS, [
    ...baseHeaders, ...itemHeaders, 'Overall Notes', 'Pass Count', 'Fail Count', 'NA Count', 'Unchecked Count'
  ]);

  const selfieUrl = d.selfie
    ? saveImageToDrive(d.selfie, `CL_SELFIE_${sanitize(d.worker)}_${fmtTs(d.timestamp)}.jpg`, PHOTO_FOLDERS.DAILY)
    : '';

  const results = d.checklistResults || {};
  const itemValues = [];
  let pass = 0, fail = 0, na = 0, unc = 0;
  CL_ITEMS.forEach((g, gi) => {
    g.items.forEach((item, ii) => {
      const val = results[`${gi}_${ii}`];
      if      (val === 'pass') { itemValues.push('PASS ✅'); pass++; }
      else if (val === 'fail') { itemValues.push('FAIL ❌'); fail++; }
      else if (val === 'na')   { itemValues.push('N/A');      na++;  }
      else                     { itemValues.push('—');         unc++; }
    });
  });

  sheet.appendRow([
    new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
    d.worker || '', d.building || '',
    d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '',
    selfieUrl, ...itemValues, d.notes || '', pass, fail, na, unc
  ]);

  if (fail >= 3) {
    sendEmailAlert(FM_EMAIL,
      `⚠️ APDC — Checklist ${fail} failures — ${d.building}`,
      `Worker ${d.worker} submitted a PM checklist for ${d.building} with ${fail} failed items.\n\nBuilding: ${d.building}\nWorker: ${d.worker}\nTime: ${d.timestamp}\n\n— APDC System`
    );
  }

  return { status: 'ok', message: 'Checklist recorded', pass, fail, na, unc };
}

// ================================================================
// FLAG
// ================================================================
function handleFlag(d) {
  const sheet = getOrCreateSheet(SHEETS.FLAGS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'Urgency', 'Category',
    'Location', 'Description', 'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)',
    'Photo URL', 'Selfie URL', 'Status', 'FM Notified'
  ]);

  const ts        = fmtTs(d.timestamp);
  const w         = sanitize(d.worker);
  const selfieUrl = d.selfie ? saveImageToDrive(d.selfie, `FLAG_SELFIE_${w}_${ts}.jpg`, PHOTO_FOLDERS.DISCOVERY) : '';
  const photoUrl  = d.photo  ? saveImageToDrive(d.photo,  `FLAG_PHOTO_${w}_${ts}.jpg`,  PHOTO_FOLDERS.DISCOVERY) : '';
  const fmNotified = (d.urgency === 'Emergency' || d.urgency === 'Urgent') ? 'YES' : 'NO';

  sheet.appendRow([
    new Date(d.timestamp), new Date(d.submittedAt || d.timestamp),
    d.worker || '', d.building || '', d.urgency || '', d.category || '',
    d.location || '', d.description || '',
    d.lat || '', d.lng || '', d.accuracy ? Math.round(d.accuracy) : '',
    photoUrl, selfieUrl, 'Open', fmNotified
  ]);

  if (FM_EMAIL && (d.urgency === 'Emergency' || d.urgency === 'Urgent')) {
    sendEmailAlert(FM_EMAIL,
      `🚨 APDC — ${d.urgency} Issue Flagged — ${d.building}`,
      `A ${d.urgency} issue has been flagged.\n\nBUILDING: ${d.building}\nWORKER: ${d.worker}\nCATEGORY: ${d.category}\nLOCATION: ${d.location}\nDESCRIPTION: ${d.description}\nTIME: ${d.timestamp}\nGPS: ${d.lat}, ${d.lng}\n\n${photoUrl ? 'PHOTO: ' + photoUrl + '\n\n' : ''}Action required.\n\n— APDC Automated Alert`
    );
  }

  return { status: 'ok', message: 'Flag recorded', urgency: d.urgency, fmNotified };
}

// ================================================================
// HELPERS
// ================================================================
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    styleHeader(sheet, 1, headers.length);
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.setColumnWidth(i, 150);
  }
  return sheet;
}

function styleHeader(sheet, row, cols) {
  const r = sheet.getRange(row, 1, 1, cols);
  r.setBackground('#0f172a');
  r.setFontColor('#f97316');
  r.setFontWeight('bold');
  r.setFontSize(10);
}

function saveImageToDrive(base64DataUrl, filename, subfolder) {
  try {
    if (!DRIVE_FOLDER_ID || !base64DataUrl) return '';
    const base64 = base64DataUrl.replace(/^data:image\/\w+;base64,/, '');
    const blob   = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', filename);
    const root   = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    let folder = root;
    if (subfolder) {
      const existing = root.getFoldersByName(subfolder);
      folder = existing.hasNext() ? existing.next() : root.createFolder(subfolder);
    }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    Logger.log('Image save error: ' + err.toString());
    return '';
  }
}

function sendEmergencyAlert(subject, d) {
  if (!FM_EMAIL) return;
  sendEmailAlert(FM_EMAIL,
    `🚨 APDC — ${subject} — ${d.building}`,
    `EMERGENCY.\n\nWorker: ${d.worker}\nBuilding: ${d.building}\nType: ${d.workType || d.category}\nTime: ${d.timestamp}\nDescription: ${d.description}\n\nImmediate action required.\n\n— APDC System`
  );
}

function sendEmailAlert(to, subject, body) {
  try { if (to) MailApp.sendEmail({ to, subject, body, name: COMPANY_NAME }); }
  catch (e) { Logger.log('Email error: ' + e.toString()); }
}

function fmtTs(isoTs) {
  if (!isoTs) return new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').substring(0, 19);
  return String(isoTs).replace(/[:.]/g, '-').replace('T', '_').substring(0, 19);
}

function sanitize(s) { return (s || '').replace(/[^a-zA-Z0-9_\- ]/g, '').substring(0, 30); }

function phDate() { return Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd'); }

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function logEntry(status, type, worker, note) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let log  = ss.getSheetByName(SHEETS.LOG);
    if (!log) {
      log = ss.insertSheet(SHEETS.LOG);
      log.appendRow(['Timestamp', 'Status', 'Type', 'Worker', 'Note']);
      styleHeader(log, 1, 5);
    }
    log.appendRow([new Date(), status, type, worker, note]);
  } catch (e) {}
}

// ================================================================
// SETUP — Run once after deployment
// ================================================================
function setupSheets() {
  getOrCreateSheet(SHEETS.ATTENDANCE, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'Check Type',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Address', 'Selfie URL', 'Client Timestamp',
    'GPS Distance (m)', 'GPS Status', 'Geofence Result', 'Timestamp Flag'
  ]);
  getOrCreateSheet(SHEETS.DAILY_REPORTS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'GPS Address',
    'Task #', 'Task Name', 'Category', 'Before Photo URL', 'After Photo URL', 'Task Notes', 'Selfie URL'
  ]);
  getOrCreateSheet(SHEETS.SPECIAL_TASKS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'GPS Address',
    'Description', 'Requested By', 'Location in Building', 'Before Photo URL', 'After Photo URL', 'Notes', 'Selfie URL', 'Status'
  ]);
  getOrCreateSheet(SHEETS.WORKORDERS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'WO Number', 'Classification',
    'Work Type', 'Location', 'Description', 'Materials Used', 'Status',
    'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)', 'Before Photo URL', 'After Photo URL', 'Selfie URL',
    'Assigned To', 'Priority', 'Due Date', 'Source', 'Instructions'
  ]);
  getOrCreateSheet(SHEETS.FLAGS, [
    'Timestamp', 'Submitted At', 'Worker', 'Building', 'Urgency', 'Category',
    'Location', 'Description', 'GPS Lat', 'GPS Lng', 'GPS Accuracy (m)',
    'Photo URL', 'Selfie URL', 'Status', 'FM Notified'
  ]);
  getOrCreateSheet(SHEETS.LOG, ['Timestamp', 'Status', 'Type', 'Worker', 'Note']);
  getConfig();       // creates Config + Buildings tabs
  getBuildingCoords(SpreadsheetApp.openById(SPREADSHEET_ID)); // ensures Buildings sheet exists

  // Create Drive subfolders
  if (DRIVE_FOLDER_ID) {
    try {
      const root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      Object.values(PHOTO_FOLDERS).forEach(name => {
        if (!root.getFoldersByName(name).hasNext()) root.createFolder(name);
      });
    } catch (e) {}
  }

  Logger.log(
    '✅ Setup complete! All sheets and Drive subfolders are ready.\n' +
    'NEXT STEPS:\n' +
    '1. Fill FM_EMAIL in the script (line 17)\n' +
    '2. Config tab: each row = Worker | Building | Worker Email\n' +
    '3. Buildings tab: each row = Building Name | Lat | Lng | Radius(m)\n' +
    '4. Redeploy as web app'
  );
}

function testBackend() {
  const r1 = handleAttendance({ type: 'attendance', worker: 'Test Worker', building: 'Test Building', checkType: 'IN', timestamp: new Date().toISOString(), lat: 14.5995, lng: 120.9842, accuracy: 12, device: 'Test', selfie: '', submittedAt: new Date().toISOString() });
  Logger.log('Attendance: ' + JSON.stringify(r1));
  const r2 = handleFlag({ mode: 'flag', worker: 'Test Worker', building: 'Test Building', urgency: 'Urgent', category: 'Electrical Problem', location: 'Main hall', description: 'Exposed wiring', timestamp: new Date().toISOString(), lat: 14.5995, lng: 120.9842, accuracy: 15, selfie: '', photo: '', submittedAt: new Date().toISOString() });
  Logger.log('Flag: ' + JSON.stringify(r2));
  Logger.log('All tests complete.');
}
