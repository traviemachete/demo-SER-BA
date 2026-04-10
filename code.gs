// ============================================================
// CM Service Web App - Google Apps Script (Server Side)
// Code.gs  –  อัปเดต: ตรง Header ของ Sheet จริง
// ============================================================

const SPREADSHEET_ID = '10SVuCx9hwj91C2B8i9kWtXry-CURkDghy4QvxW7Hh7M';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BIOAXEL CM Service – ระบบแจ้งซ่อม')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── Helper ──────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) Logger.log('Sheet not found: ' + name);
  return sheet;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = r[i];
        // ← เพิ่มบรรทัดนี้: แปลง Date object → string ก่อน return
        if (val instanceof Date) {
          val = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
        }
        obj[String(h).trim()] = val;
      });
      return obj;
    });
}

// ── Main Project ─────────────────────────────────────────────
// Sheet: Main_Project
// normalize column names เพราะ Sheet อาจใช้ชื่อต่างกัน
function getProjects() {
  try {
    const rows = sheetToObjects(getSheet('Main_Project'));
    return rows.map(r => ({
      Contract_No:  String(r['Contract_No']  || '').trim(),
      Project_Name: String(r['Project_Name'] || '').trim(),
      Type:         String(r['Type']         || '').trim(),   // e.g. BA100M-YMD-06
      Model:        String(r['Model']        || r['Product_Model'] || '').trim(),
      contact_Tel:  String(r['contact_Tel']  || r['contact_Name']  || r['Tel'] || '').trim(),
      Region:       String(r['Region']       || r['Province']      || '').trim(),
    })).filter(r => r.Contract_No !== '' || r.Type !== '');
  } catch(e) {
    Logger.log('getProjects error: ' + e.message);
    return [];
  }
}

// ── Asset Master ─────────────────────────────────────────────
function getAssets() {
  try {
    return sheetToObjects(getSheet('Asset_Master'));
  } catch(e) {
    Logger.log('getAssets error: ' + e.message);
    return [];
  }
}

// ── Knowledge Base ───────────────────────────────────────────
function getKnowledgeByAsset(assetCode) {
  try {
    const data = sheetToObjects(getSheet('Knowledge_Base'));
    if (!assetCode || assetCode === '') return data;
    return data.filter(r => {
      const code = String(r['Asset_Code'] || '').trim();
      return code === assetCode || code === 'ALL' || code === '';
    });
  } catch(e) {
    Logger.log('getKnowledgeByAsset error: ' + e.message);
    return [];
  }
}

// ── Spare Parts ──────────────────────────────────────────────
function getSpareParts() {
  try {
    return sheetToObjects(getSheet('Spare_Parts'));
  } catch(e) {
    Logger.log('getSpareParts error: ' + e.message);
    return [];
  }
}

// ── Service Records ──────────────────────────────────────────
// Header จริงในชีท:
// Ticket_ID | Report_Date | Contract_No | Product_Model | Project_Name |
// Warranty_Status | Asset_Code | Asset_Name | Client_Issue | Technical_Issue |
// Root_Cause | Action_Taken | Spare_Parts_Used | Service_Date | Finish_Date |
// Technician | Work_Status | SLA_Days | Cost | Photo_Before | Photo_After | Remark

function getServiceRecords() {
  try {
    return sheetToObjects(getSheet('Service_Records'));
  } catch(e) {
    return [];
  }
}

function saveServiceRecord(record) {
  try {
    const sheet = getSheet('Service_Records');
    if (!sheet) throw new Error('ไม่พบชีท Service_Records');

    // Header ตรงกับ Sheet จริง
    const HEADERS = [
      'Ticket_ID','Report_Date','Contract_No','Product_Model','Project_Name',
      'Warranty_Status','Asset_Code','Asset_Name','Client_Issue','Technical_Issue',
      'Root_Cause','Action_Taken','Spare_Parts_Used','Service_Date','Finish_Date',
      'Technician','Work_Status','SLA_Days','Cost','Photo_Before','Photo_After','Remark'
    ];

    // ถ้า Sheet ว่างให้ใส่ Header ก่อน
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }

    // อ่าน Header จริงจาก Sheet row 1
    const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
    const liveHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                          .map(h => String(h).trim())
                          .filter(h => h !== '');

    // สร้าง Ticket ID
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMdd');
    const seq = String(sheet.getLastRow()).padStart(4, '0');
    const ticketId = 'TK-' + dateStr + '-' + seq;

    // Map form fields → sheet headers
    // form field           sheet header
    const normalized = {
      'Ticket_ID':      ticketId,
      'Report_Date':    Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
      'Service_Date':   Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
      'Contract_No':    record['Contract_No']    || '',
      'Project_Name':   record['Project_Name']   || '',
      'Product_Model':  record['Model']          || '',   // form=Model → sheet=Product_Model
      'Warranty_Status':record['wrt_Status']     || '',   // form=wrt_Status → Warranty_Status
      'Asset_Code':     record['Asset_Code']     || '',
      'Asset_Name':     record['Asset_Name']     || '',
      'Client_Issue':   record['Problem_Detail'] || '',   // form=Problem_Detail → Client_Issue
      'Technical_Issue':record['mac_Status']     || '',   // mac_Status → Technical_Issue
      'Action_Taken':   record['Action_Taken']   || '',
      'Technician':     record['Technician']     || '',
      'Work_Status':    record['Work_Status']    || '',
      'Remark':         record['Remark']         || '',
    };

    const row = liveHeaders.map(h => {
      const val = normalized[h];
      return (val !== undefined && val !== null) ? val : '';
    });

    sheet.appendRow(row);
    try { sheet.autoResizeColumns(1, liveHeaders.length); } catch(e2) {}

    return { success: true, ticketId: ticketId };
  } catch (e) {
    Logger.log('saveServiceRecord error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Dashboard Summary ────────────────────────────────────────
function getDashboardData() {
  try {
    const tickets = sheetToObjects(getSheet('Service_Records'));

    const statusCount  = {};
    const regionCount  = {};
    const modelCount   = {};
    const monthlyCount = {};
    const symptomCount = {};
    let   doneFirst    = 0; // ซ่อมสำเร็จครั้งแรก (ไม่มี Root_Cause = ต้องส่งซ่อมซ้ำ)
    const slaBreached  = [];

    tickets.forEach(r => {
      const s  = String(r['Work_Status']   || 'ไม่ระบุ').trim();
      const rg = String(r['Region']        || 'ไม่ระบุ').trim();
      const m  = String(r['Product_Model'] || r['Model'] || 'ไม่ระบุ').trim();
      const ts = r['Report_Date'] ? String(r['Report_Date']).substring(0, 7) : 'ไม่ระบุ';
      const sym = String(r['Client_Issue'] || '').trim();

      statusCount[s]   = (statusCount[s]   || 0) + 1;
      regionCount[rg]  = (regionCount[rg]  || 0) + 1;
      modelCount[m]    = (modelCount[m]    || 0) + 1;
      monthlyCount[ts] = (monthlyCount[ts] || 0) + 1;
      if (sym) symptomCount[sym] = (symptomCount[sym] || 0) + 1;

      // First-time fix: Work_Status = ซ่อมเสร็จสิ้น
      if (s === 'ซ่อมเสร็จสิ้น') doneFirst++;

      // SLA: SLA_Days มีค่า และ Work_Status ไม่ใช่เสร็จสิ้น
      const sla = parseInt(r['SLA_Days']) || 0;
      if (sla > 0 && s !== 'ซ่อมเสร็จสิ้น') {
        slaBreached.push({
          Ticket_ID:    r['Ticket_ID']    || '',
          Asset_Code:   r['Asset_Code']   || '',
          Work_Status:  s,
          SLA_Days:     sla,
          Report_Date:  r['Report_Date']  || '',
          Client_Issue: sym
        });
      }
    });

    // Top 5 symptoms
    const top5Symptoms = Object.entries(symptomCount)
      .sort((a,b) => b[1]-a[1]).slice(0,5)
      .map(([name, count]) => ({ name, count }));

    // First-time fix rate
    const totalDone  = statusCount['ซ่อมเสร็จสิ้น'] || 0;
    const fixRate    = totalDone > 0 ? Math.round((doneFirst / totalDone) * 100) : 0;

    return {
      total: tickets.length,
      statusCount, regionCount, modelCount, monthlyCount,
      top5Symptoms, fixRate,
      slaBreached: slaBreached.slice(0, 20) // max 20 rows
    };
  } catch(e) {
    Logger.log('getDashboardData error: ' + e.message);
    return { total: 0, statusCount: {}, regionCount: {}, modelCount: {}, monthlyCount: {}, top5Symptoms: [], fixRate: 0, slaBreached: [] };
  }
}

// ── Ticket History ───────────────────────────────────────────
// ดึง Ticket ทั้งหมด พร้อม filter ฝั่ง server
function getTickets(filter) {
  try {
    const sheet = getSheet('Service_Records');
    if (!sheet) return [];                          // ← เพิ่ม: Sheet ไม่เจอให้ return []
    
    const all = sheetToObjects(sheet);
    Logger.log('getTickets: total=' + all.length + ' filter=' + JSON.stringify(filter));
    
    if (!filter || Object.keys(filter).length === 0) return all;  // ← ไม่มี filter คืนทั้งหมด
    
    return all.filter(r => {
      if (filter.type) {
        const clean = s => (s||'').replace(/-/g,'').toUpperCase();
        if (!clean(r['Product_Model']).includes(clean(filter.type))) return false;
      }
      if (filter.status && (r['Work_Status']||'') !== filter.status) return false;
      if (filter.q) {
        const q = filter.q.toLowerCase();
        const s = [r['Ticket_ID'],r['Client_Issue'],r['Action_Taken'],r['Technician'],r['Asset_Code']]
                  .join(' ').toLowerCase();
        if (!s.includes(q)) return false;
      }
      return true;
    });
  } catch(e) {
    Logger.log('getTickets error: ' + e.message);
    return [];   // ← return [] แทน throw เพื่อให้ client ไม่ crash
  }
}

// ── Debug helpers (รันใน GAS Editor → View Logs) ─────────────
function listSheetNames() {
  const names = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().map(s => s.getName());
  Logger.log('All sheets: ' + JSON.stringify(names));
}

function testAllFunctions() {
  Logger.log('Projects: '  + JSON.stringify(getProjects()).substring(0, 300));
  Logger.log('Assets: '    + JSON.stringify(getAssets()).substring(0, 300));
  Logger.log('Knowledge: ' + JSON.stringify(getKnowledgeByAsset('')).substring(0, 300));
  Logger.log('Parts: '     + JSON.stringify(getSpareParts()).substring(0, 300));
  Logger.log('Dashboard: ' + JSON.stringify(getDashboardData()));
}

function debugTickets() {
  const sheet = getSheet('Service_Records');
  Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));
  if (!sheet) return;
  Logger.log('Last row: ' + sheet.getLastRow());
  const tickets = sheetToObjects(sheet);
  Logger.log('Total objects: ' + tickets.length);
  if (tickets.length > 0) Logger.log('Sample: ' + JSON.stringify(tickets[0]));
}
