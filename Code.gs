// ============================================================
//  LINK HUB — Google Apps Script (Backend API)
//  ใช้ได้กับทั้ง ruangchai_links และ template_links
//
//  วิธีติดตั้ง:
//  1. เปิด Google Sheet ใหม่
//  2. Extensions → Apps Script
//  3. ลบโค้ดเดิม วางโค้ดนี้ทั้งหมด
//  4. กด Deploy → New deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  5. คัดลอก URL ที่ได้ไปวางในไฟล์ HTML (ตรง API_URL)
// ============================================================

const SHEET_NAME  = 'Links';
const PASS_USER   = '1234';       // ← รหัสผู้ใช้ทั่วไป
const PASS_ADMIN  = 'admin1234';  // ← รหัส Admin

function doGet(e) {
  // กัน error เมื่อ Run ตรงใน Editor โดยไม่มี event
  if (!e || !e.parameter) return json(getLinks());
  const p = e.parameter;
  if (p.action === 'getSettings') return json(getSettings());
  return json(getLinks());
}

function doPost(e) {
  if (!e || !e.postData) return json({ error: 'No data' });
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return json({ error: 'Invalid JSON' }); }

  if (!body || body.adminToken !== PASS_ADMIN)
    return json({ error: 'Unauthorized' });

  const a = body.action;
  if (a === 'add')          return json(addLink(body.data));
  if (a === 'update')       return json(updateLink(body.id, body.data));
  if (a === 'delete')       return json(deleteLink(body.id));
  if (a === 'saveSettings') return json(saveSettings(body.data));
  if (a === 'seed')         return json(seedLinks(body.links));
  if (a === 'clearAll')     return json(clearAll());
  return json({ error: 'Unknown action' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id','name','url','cat','note']);
    sh.getRange(1,1,1,5).setFontWeight('bold');
  }
  return sh;
}

function getLinks() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { links: [] };
  const links = data.slice(1)
    .filter(r => r[0] !== '')
    .map(r => ({ id:String(r[0]),name:String(r[1]),url:String(r[2]),cat:String(r[3]),note:String(r[4]) }));
  return { links };
}

function addLink(d) {
  const id = Date.now().toString();
  getSheet().appendRow([id, d.name||'', d.url||'', d.cat||'', d.note||'']);
  return { success:true, id };
}

function updateLink(id, d) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sh.getRange(i+1,1,1,5).setValues([[
        id,
        d.name  !== undefined ? d.name  : rows[i][1],
        d.url   !== undefined ? d.url   : rows[i][2],
        d.cat   !== undefined ? d.cat   : rows[i][3],
        d.note  !== undefined ? d.note  : rows[i][4]
      ]]);
      return { success:true };
    }
  }
  return { error:'Not found' };
}

function deleteLink(id) {
  const sh = getSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sh.deleteRow(i+1);
      return { success:true };
    }
  }
  return { error:'Not found' };
}

function seedLinks(links) {
  const sh = getSheet();
  links.forEach((d,i) => {
    sh.appendRow([d.id||(Date.now()+i).toString(), d.name||'', d.url||'', d.cat||'', d.note||'']);
  });
  return { success:true, count:links.length };
}

function clearAll() {
  // ลบข้อมูลทุก row ใน Sheet (เหลือแค่ header)
  const sh = getSheet();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow - 1);

  // ลบ Settings ทั้งหมด
  PropertiesService.getScriptProperties().deleteAllProperties();

  return { success: true };
}


function getSettings() {
  const p = PropertiesService.getScriptProperties();
  return { settings:{
    badge:  p.getProperty('badge')  || 'หน่วยงาน',
    title:  p.getProperty('title')  || 'รวมลิงก์ระบบงาน',
    sub:    p.getProperty('sub')    || 'คำอธิบาย',
    logo:   p.getProperty('logo')   || 'LOGO',
    footer: p.getProperty('footer') || 'รวมลิงก์ระบบงาน',
    accent: p.getProperty('accent') || '#c9963a,#e8b84b'
  }};
}

function saveSettings(d) {
  const p = PropertiesService.getScriptProperties();
  ['badge','title','sub','logo','footer','accent'].forEach(k => {
    if (d[k] !== undefined) p.setProperty(k, d[k]);
  });
  return { success:true };
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
