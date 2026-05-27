// ============================================================
// เว็บโรงเรียน - Code.gs (Server-side)
// ============================================================

const CONFIG = {
  SHEET_ID: '1Lf_s0heJKdKZU8TSZcklqnSMw5I8_gL3VhSvPUrPY08',
  DRIVE_FOLDER_ID: '1u-g366a1GCG1DM_s6bXTsmCVRA5DCA54',
  SCHOOL_NAME: 'โรงเรียนตัวอย่าง',
  SESSION_HOURS: 8
};

// ===== HELPER: DYNAMIC DB & DRIVE RESOLVER =====
function getActiveSS() {
  let ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}
  
  if (CONFIG.SHEET_ID && CONFIG.SHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE') {
    try {
      return SpreadsheetApp.openById(CONFIG.SHEET_ID);
    } catch (err) {
      throw new Error("เชื่อมต่อ Spreadsheet ตาม ID ที่ระบุล้มเหลว: " + err.message);
    }
  }
  
  throw new Error("ไม่สามารถเชื่อมต่อฐานข้อมูลได้: กรุณาเปิดใช้งานสคริปต์นี้จาก Google Sheets หรือตั้งค่า SHEET_ID ใน CONFIG");
}

function getUploadFolder() {
  let folderId = '';
  if (CONFIG.DRIVE_FOLDER_ID && CONFIG.DRIVE_FOLDER_ID !== 'YOUR_DRIVE_FOLDER_ID_HERE') {
    folderId = CONFIG.DRIVE_FOLDER_ID;
  } else {
    folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    if (!folderId) {
      try {
        const folder = DriveApp.createFolder(CONFIG.SCHOOL_NAME + ' - Images');
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        folderId = folder.getId();
        PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folderId);
      } catch (e) {
        throw new Error("สร้างโฟลเดอร์สำหรับเก็บภาพล้มเหลว: " + e.message);
      }
    }
  }
  return DriveApp.getFolderById(folderId);
}

// ===== HELPER: SANITIZE DATE OBJECTS FOR JSON SERIALIZATION =====
function sanitizeForJson(val) {
  if (val instanceof Date) {
    try {
      return val.toISOString();
    } catch (e) {
      return val.toString();
    }
  }
  if (Array.isArray(val)) {
    return val.map(sanitizeForJson);
  }
  if (val !== null && typeof val === 'object') {
    const copy = {};
    for (const key in val) {
      if (val.hasOwnProperty(key)) {
        copy[key] = sanitizeForJson(val[key]);
      }
    }
    return copy;
  }
  return val;
}

// ===== ROUTING & AUTO INIT =====
function doGet(e) {
  // Auto-init DB tables if first deployment run
  try {
    const ss = getActiveSS();
    if (!ss.getSheetByName('Settings') || !ss.getSheetByName('Admins') || !ss.getSheetByName('Board') || !ss.getSheetByName('Messages') || !ss.getSheetByName('Updates')) {
      setupSpreadsheet();
    }
  } catch (err) {}

  let title = CONFIG.SCHOOL_NAME;
  try {
    const settings = getSettings();
    if (settings && settings.siteTitle) {
      title = settings.siteTitle;
    } else if (settings && settings.schoolName) {
      title = settings.schoolName;
    }
  } catch (err) {}

  const tmpl = HtmlService.createTemplateFromFile('index');
  return tmpl.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== BATCH DATA RETRIEVAL =====
function getPublicData() {
  return {
    settings: getSettings(),
    news: getNews(true),
    activities: getActivities(true),
    events: getEvents(true),
    board: getBoard(true),
    updates: getUpdates(true)
  };
}

// ===== SETUP DATABASE (CMS Tables Initialization) =====
function setupSpreadsheet() {
  const ss = getActiveSS();

  // Sheet: News
  if (!ss.getSheetByName('News')) {
    const s = ss.insertSheet('News');
    s.appendRow(['id','title','content','imageUrl','category','author','date','published']);
    s.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white');
    s.setColumnWidth(3, 400);
  }

  // Sheet: Activities
  if (!ss.getSheetByName('Activities')) {
    const s = ss.insertSheet('Activities');
    s.appendRow(['id','title','description','imageUrl','date','published']);
    s.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1E8449').setFontColor('white');
  }

  // Sheet: Events
  if (!ss.getSheetByName('Events')) {
    const s = ss.insertSheet('Events');
    s.appendRow(['id','title','description','eventDate','time','location','color','published']);
    s.getRange(1,1,1,8).setFontWeight('bold').setBackground('#7D3C98').setFontColor('white');
  }

  // Sheet: Settings (CMS-driven config)
  if (!ss.getSheetByName('Settings')) {
    const s = ss.insertSheet('Settings');
    s.appendRow(['key','value']);
    [
      ['schoolName','โรงเรียนตัวอย่าง'],
      ['siteTitle','โรงเรียนตัวอย่าง - ยินดีต้อนรับสู่ระบบโรงเรียนออนไลน์'],
      ['schoolSubtitle','มุ่งมั่น ซื่อสัตย์ พัฒนา'],
      ['schoolAddress','123 ถ.ตัวอย่าง อ.เมือง จ.กรุงเทพฯ 10000'],
      ['contactPhone','02-000-0000'],
      ['contactEmail','school@example.ac.th'],
      ['logoUrl',''],
      ['heroImageUrl',''],
      ['facebookUrl',''],
      ['lineOaUrl',''],
      
      // Dynamic About Us Settings
      ['aboutHistory', 'โรงเรียนมุ่งเน้นความเป็นเลิศทางด้านการศึกษาควบคู่ไปกับคุณธรรมจริยธรรม ก่อตั้งขึ้นด้วยความมุ่งมั่นที่จะพัฒนาเยาวชนให้เติบโตขึ้นเป็นบุคลากรที่มีความรู้ ความสามารถ และทักษะที่จำเป็นในศตวรรษที่ 21 ภายใต้สภาพแวดล้อมที่ส่งเสริมการเรียนรู้และการฝึกฝนทักษะรอบด้าน'],
      ['aboutVision', '"เป็นเลิศในวิชาการ พัฒนาทักษะชีวิต เพียบพร้อมด้วยคุณธรรม นำชุมชนสู่วิถีการเรียนรู้ยั่งยืน"'],
      ['aboutMission', '1. จัดกระบวนการเรียนรู้ที่สอดคล้องกับมาตรฐานการศึกษายุคดิจิทัล\n2. ปลูกฝังค่านิยมด้านความซื่อสัตย์ มีวินัย และรักษ์ความเป็นไทย\n3. พัฒนาทักษะอาชีพและทักษะชีวิตเพื่ออนาคตของผู้เรียน'],
      ['aboutDirectorName', 'ผอ.สมชาย ใจดี'],
      ['aboutDirectorRole', 'ผู้อำนวยการโรงเรียน'],
      ['aboutDirectorMessage', 'ยินดีต้อนรับสู่ระบบโรงเรียนออนไลน์ เรามุ่งหวังที่จะพัฒนาเยาวชนและเตรียมความพร้อมสู่ศตวรรษที่ 21 ร่วมมือร่วมใจกันสร้างสังคมแห่งการเรียนรู้สืบไป'],
      ['aboutDirectorImageUrl', ''],
      ['aboutDeputy1Name', 'รองฯ สมศรี รักเรียน'],
      ['aboutDeputy1Role', 'รองผู้อำนวยการฝ่ายวิชาการ'],
      ['aboutDeputy1ImageUrl', ''],
      ['aboutDeputy2Name', 'รองฯ สมศักดิ์ มุ่งมั่น'],
      ['aboutDeputy2Role', 'รองผู้อำนวยการฝ่ายบริหารงานทั่วไป'],
      ['aboutDeputy2ImageUrl', ''],
      ['contactMapEmbedUrl', 'https://www.openstreetmap.org/export/embed.html?bbox=100.485%2C13.735%2C100.515%2C13.765&layer=mapnik'],
      ['geminiApiKey', ''],
      ['groqApiKey', ''],
      ['aiTextProvider', 'gemini'],
      
      // Design & Theme Customization (Appearance Settings)
      ['themeColorPrimary', '#0F172A'],     // Slate 900
      ['themeColorSecondary', '#1E293B'],   // Slate 800
      ['themeColorGold', '#D4AF37'],        // Gold Accent
      ['themeBackgroundColor', '#F8FAFC'],  // Slate 50
      ['themeFontFamily', 'Sarabun'],
      ['themeFontSizeBase', '16px'],
      ['themeFontSizeHeading', '1.85rem']
    ].forEach(r => s.appendRow(r));
    s.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white');
  }

  // Sheet: Admins
  if (!ss.getSheetByName('Admins')) {
    const s = ss.insertSheet('Admins');
    s.appendRow(['username','password_hash','name','created_at']);
    s.appendRow(['admin', hashPassword('admin1234'), 'ผู้ดูแลระบบ', new Date()]);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#922B21').setFontColor('white');
  }

  // Sheet: Board (Dynamic executives list)
  if (!ss.getSheetByName('Board')) {
    const s = ss.insertSheet('Board');
    s.appendRow(['id','name','role','imageUrl','published','order']);
    s.getRange(1,1,1,6).setFontWeight('bold').setBackground('#7D6608').setFontColor('white');
    
    // Attempt migration from existing settings to prevent data loss
    let dirName = 'ผอ.สมชาย ใจดี', dirRole = 'ผู้อำนวยการโรงเรียน', dirImg = '';
    let dep1Name = 'รองฯ สมศรี รักเรียน', dep1Role = 'รองผู้อำนวยการฝ่ายวิชาการ', dep1Img = '';
    let dep2Name = 'รองฯ สมศักดิ์ มุ่งมั่น', dep2Role = 'รองผู้อำนวยการฝ่ายบริหารงานทั่วไป', dep2Img = '';
    
    try {
      const settingsSheet = ss.getSheetByName('Settings');
      if (settingsSheet) {
        const data = settingsSheet.getDataRange().getValues();
        data.forEach(r => {
          if (r[0] === 'aboutDirectorName') dirName = r[1];
          if (r[0] === 'aboutDirectorRole') dirRole = r[1];
          if (r[0] === 'aboutDirectorImageUrl') dirImg = r[1];
          if (r[0] === 'aboutDeputy1Name') dep1Name = r[1];
          if (r[0] === 'aboutDeputy1Role') dep1Role = r[1];
          if (r[0] === 'aboutDeputy1ImageUrl') dep1Img = r[1];
          if (r[0] === 'aboutDeputy2Name') dep2Name = r[1];
          if (r[0] === 'aboutDeputy2Role') dep2Role = r[1];
          if (r[0] === 'aboutDeputy2ImageUrl') dep2Img = r[1];
        });
      }
    } catch (e) {}

    s.appendRow([Utilities.getUuid(), dirName, dirRole, dirImg, true, 1]);
    s.appendRow([Utilities.getUuid(), dep1Name, dep1Role, dep1Img, true, 2]);
    s.appendRow([Utilities.getUuid(), dep2Name, dep2Role, dep2Img, true, 3]);
  }

  // Sheet: Messages (Contact forms inbox)
  if (!ss.getSheetByName('Messages')) {
    const s = ss.insertSheet('Messages');
    s.appendRow(['id','name','email','subject','message','date','status','notes']);
    s.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2E4053').setFontColor('white');
  }

  // Sheet: Updates (Announcements and blog posts)
  if (!ss.getSheetByName('Updates')) {
    const s = ss.insertSheet('Updates');
    s.appendRow(['id','title','content','date','published']);
    s.getRange(1,1,1,5).setFontWeight('bold').setBackground('#2C3E50').setFontColor('white');
  }

  // Migration for existing databases (ensure geminiApiKey, groqApiKey, aiTextProvider setting keys exist)
  try {
    const settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      let hasGeminiKey = false;
      let hasGroqKey = false;
      let hasAiProvider = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'geminiApiKey') hasGeminiKey = true;
        if (data[i][0] === 'groqApiKey') hasGroqKey = true;
        if (data[i][0] === 'aiTextProvider') hasAiProvider = true;
      }
      if (!hasGeminiKey) settingsSheet.appendRow(['geminiApiKey', '']);
      if (!hasGroqKey) settingsSheet.appendRow(['groqApiKey', '']);
      if (!hasAiProvider) settingsSheet.appendRow(['aiTextProvider', 'gemini']);
    }
  } catch (e) {}

  return '✅ Setup สำเร็จ! เปิด Web App ได้เลย';
}

// ===== AUTH =====
function login(username, password) {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Admins');
    if (!sheet) return { success: false, message: 'ไม่พบข้อมูลผู้ดูแลระบบ' };

    const data = sheet.getDataRange().getValues();
    const hashed = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim() &&
          String(data[i][1]).trim() === hashed) {
        const token = Utilities.getUuid();
        const expiry = Date.now() + CONFIG.SESSION_HOURS * 3600000;
        PropertiesService.getScriptProperties().setProperty(
          'SESSION_' + token,
          JSON.stringify({ username, name: data[i][2], expiry })
        );
        return { success: true, token, name: data[i][2] };
      }
    }
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== LOGOUT =====
function logout(token) {
  PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
  return { success: true };
}

// ===== VALIDATE SESSION =====
function validateSession(token) {
  if (!token) return false;
  const raw = PropertiesService.getScriptProperties().getProperty('SESSION_' + token);
  if (!raw) return false;
  const session = JSON.parse(raw);
  if (Date.now() > session.expiry) {
    PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
    return false;
  }
  return true;
}

// ===== CHANGE PASSWORD =====
function changePassword(token, oldPw, newPw) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const raw = PropertiesService.getScriptProperties().getProperty('SESSION_' + token);
  const { username } = JSON.parse(raw);
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('Admins');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (data[i][1] !== hashPassword(oldPw)) return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      sheet.getRange(i + 1, 2).setValue(hashPassword(newPw));
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้' };
}

// ===== HASH PASSWORD =====
function hashPassword(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ===== CMS: SETTINGS SYSTEM =====
function getSettings() {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues().slice(1);
    const obj = {};
    data.forEach(r => { obj[r[0]] = r[1]; });
    return sanitizeForJson(obj);
  } catch (e) { return {}; }
}

// ===== HELPER: PARSE & RESOLVE GOOGLE MAPS / OSM URLS =====
function processMapUrl(url) {
  if (!url) return '';
  url = url.trim();

  // 1. Extract src if iframe code is provided
  if (url.toLowerCase().indexOf('<iframe') !== -1) {
    const match = url.match(/src=["']([^"']+)["']/i);
    if (match && match[1]) {
      url = match[1];
    }
  }

  url = url.trim();

  // Auto-prefix protocol if missing
  if (url && url.indexOf('http') !== 0) {
    if (url.indexOf('maps.app.goo.gl') === 0 || 
        url.indexOf('goo.gl/maps') === 0 || 
        url.indexOf('google.com') === 0 || 
        url.indexOf('maps.google') === 0 || 
        url.indexOf('openstreetmap.org') === 0 ||
        url.indexOf('www.openstreetmap.org') === 0) {
      url = 'https://' + url;
    }
  }

  // 2. Return immediately if it's already an embed URL
  if (url.indexOf('/maps/embed') !== -1 || url.indexOf('output=embed') !== -1 || url.indexOf('openstreetmap.org/export/embed.html') !== -1) {
    return url;
  }

  // 3. Resolve redirect for Google Maps short links (maps.app.goo.gl, goo.gl/maps)
  if (url.indexOf('maps.app.goo.gl') !== -1 || url.indexOf('goo.gl/maps') !== -1) {
    try {
      let currentUrl = url;
      for (let i = 0; i < 3; i++) {
        const response = UrlFetchApp.fetch(currentUrl, {
          followRedirects: false,
          muteHttpExceptions: true
        });
        const code = response.getResponseCode();
        if (code >= 300 && code < 400) {
          const headers = response.getHeaders();
          const redirectUrl = headers['location'] || headers['Location'];
          if (redirectUrl) {
            if (redirectUrl.indexOf('/') === 0) {
              const originMatch = currentUrl.match(/^(https?:\/\/[^/]+)/);
              currentUrl = (originMatch ? originMatch[1] : '') + redirectUrl;
            } else {
              currentUrl = redirectUrl;
            }
          } else {
            break;
          }
        } else {
          break;
        }
      }
      url = currentUrl;
    } catch (e) {
      Logger.log('Error resolving Google Maps redirect: ' + e.toString());
    }
  }

  // 4. Try extracting latitude/longitude coordinates
  let latLngMatch = url.match(/\/search\/([-+]?[0-9.]+),\s*\+?([-+]?[0-9.]+)/i) || 
                     url.match(/\/place\/([-+]?[0-9.]+),\s*\+?([-+]?[0-9.]+)/i) ||
                     url.match(/@([-+]?[0-9.]+),([-+]?[0-9.]+)/i) ||
                     url.match(/[?&]q=([-+]?[0-9.]+),([-+]?[0-9.]+)/i);

  if (latLngMatch && latLngMatch[1] && latLngMatch[2]) {
    return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + latLngMatch[1] + ',' + latLngMatch[2] + '!6i17';
  }

  // 5. Try extracting search text query
  let queryMatch = url.match(/\/place\/([^/]+)/i) || url.match(/[?&]q=([^&]+)/i);
  if (queryMatch && queryMatch[1]) {
    return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + queryMatch[1] + '!6i17';
  }

  // 6. Check if input is a raw lat/long string, e.g. "16.429393, 102.823264"
  const plainCoords = url.match(/^\s*([-+]?[0-9.]+)\s*,\s*([-+]?[0-9.]+)\s*$/);
  if (plainCoords) {
    return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + plainCoords[1] + ',' + plainCoords[2] + '!6i17';
  }

  // 7. Standard web URLs that are not embedded yet
  if (url.indexOf('http') === 0) {
    if (url.indexOf('google.com/maps') !== -1 || url.indexOf('maps.google.com') !== -1) {
      const qMatch = url.match(/[?&]q=([^&]+)/i);
      if (qMatch && qMatch[1]) {
        return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + qMatch[1] + '!6i17';
      }
      return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + encodeURIComponent(url) + '!6i17';
    }
    return url;
  }

  // 8. Otherwise treat input as plain text address query
  return 'https://www.google.com/maps/embed?origin=mfe&pb=!1m3!2m1!1s' + encodeURIComponent(url) + '!6i17';
}

function updateSettings(token, settings) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  Object.keys(settings).forEach(key => {
    let value = settings[key];
    if (key === 'contactMapEmbedUrl') {
      value = processMapUrl(value);
    }
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([key, value]);
  });
  return { success: true };
}

// ===== CMS: NEWS CRUD =====
function getNews(publishedOnly) {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('News');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    if (publishedOnly) rows = rows.filter(r => r.published === true || r.published === 'TRUE');
    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    return sanitizeForJson(rows);
  } catch (e) { return []; }
}

function addNews(token, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('News');
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    data.title || '',
    data.content || '',
    data.imageUrl || '',
    data.category || 'ทั่วไป',
    data.author || 'Admin',
    new Date().toISOString(),
    true
  ]);
  return { success: true, id };
}

function updateNews(token, id, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('News');
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (all[i][0] === id) {
      const r = i + 1;
      const map = { title:2, content:3, imageUrl:4, category:5, author:6, published:8 };
      Object.keys(data).forEach(k => {
        if (map[k]) sheet.getRange(r, map[k]).setValue(data[k]);
      });
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function deleteNews(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('News');
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (all[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ===== CMS: ACTIVITIES CRUD =====
function getActivities(publishedOnly) {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Activities');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    if (publishedOnly) rows = rows.filter(r => r.published === true || r.published === 'TRUE');
    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    return sanitizeForJson(rows);
  } catch (e) { return []; }
}

function addActivity(token, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('Activities');
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    data.title || '',
    data.description || '',
    data.imageUrl || '',
    new Date().toISOString(),
    true
  ]);
  return { success: true, id };
}

function updateActivity(token, id, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('Activities');
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (all[i][0] === id) {
      const r = i + 1;
      const map = { title:2, description:3, imageUrl:4, published:6 };
      Object.keys(data).forEach(k => {
        if (map[k]) sheet.getRange(r, map[k]).setValue(data[k]);
      });
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function deleteActivity(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  const ss = getActiveSS();
  const sheet = ss.getSheetByName('Activities');
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (all[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ===== CMS: EVENTS (CALENDAR) CRUD =====
function getEventsSheet() {
  const ss = getActiveSS();
  let sheet = ss.getSheetByName('Events');
  if (!sheet) {
    setupSpreadsheet();
    sheet = ss.getSheetByName('Events');
  }
  return sheet;
}

function getEvents(publishedOnly) {
  try {
    const sheet = getEventsSheet();
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    if (publishedOnly) rows = rows.filter(r => r.published === true || r.published === 'TRUE');
    
    // Sort by eventDate ascending (earlier events first)
    rows.sort((a, b) => new Date(a.eventDate) - new Date(b.eventDate));
    return sanitizeForJson(rows);
  } catch (e) { return []; }
}

function addEvent(token, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const sheet = getEventsSheet();
    if (!sheet) return { success: false, message: 'ไม่สามารถสร้างตาราง Events ได้' };
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      data.title || '',
      data.description || '',
      data.eventDate || '',
      data.time || '',
      data.location || '',
      data.color || 'blue',
      true // published
    ]);
    return { success: true, id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateEvent(token, id, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const sheet = getEventsSheet();
    if (!sheet) return { success: false, message: 'ไม่พบตาราง Events' };
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        const r = i + 1;
        const map = { title:2, description:3, eventDate:4, time:5, location:6, color:7, published:8 };
        Object.keys(data).forEach(k => {
          if (map[k]) sheet.getRange(r, map[k]).setValue(data[k]);
        });
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteEvent(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const sheet = getEventsSheet();
    if (!sheet) return { success: false, message: 'ไม่พบตาราง Events' };
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== IMAGE AUTO UPLOAD ENGINE =====
function uploadImage(token, base64Data, filename, mimeType) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const parentFolder = getUploadFolder();
    
    // Get upload date in yyyy-MM-dd format
    let dateStr = '';
    try {
      dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (e) {
      const d = new Date();
      dateStr = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
    }
    
    // Find or create subfolder for the current date
    let uploadFolder;
    const folders = parentFolder.getFoldersByName(dateStr);
    if (folders.hasNext()) {
      uploadFolder = folders.next();
    } else {
      uploadFolder = parentFolder.createFolder(dateStr);
      uploadFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data), mimeType, filename
    );
    const file = uploadFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    return {
      success: true,
      url: `https://drive.google.com/thumbnail?id=${fileId}&sz=w1200`,
      fileId
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== CMS: BOARD MEMBERS (EXECUTIVES) CRUD =====
function getBoard(publishedOnly) {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Board');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    if (publishedOnly) rows = rows.filter(r => r.published === true || r.published === 'TRUE');
    
    // Sort by order ascending
    rows.sort((a, b) => {
      const orderA = parseInt(a.order) || 999;
      const orderB = parseInt(b.order) || 999;
      return orderA - orderB;
    });
    return sanitizeForJson(rows);
  } catch (e) { return []; }
}

function addBoardMember(token, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Board');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลผู้บริหาร' };
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      data.name || '',
      data.role || '',
      data.imageUrl || '',
      data.published !== false,
      parseInt(data.order) || (sheet.getLastRow())
    ]);
    return { success: true, id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateBoardMember(token, id, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Board');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลผู้บริหาร' };
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        const r = i + 1;
        const map = { name: 2, role: 3, imageUrl: 4, published: 5, order: 6 };
        Object.keys(data).forEach(k => {
          if (map[k] !== undefined) sheet.getRange(r, map[k]).setValue(data[k]);
        });
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลผู้บริหาร' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteBoardMember(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Board');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลผู้บริหาร' };
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'ไม่พบข้อมูล' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ===== CMS: CONTACT MESSAGES SYSTEM =====
function addContactMessage(data) {
  try {
    if (!data.name || !data.email || !data.subject || !data.message) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน' };
    }

    const ss = getActiveSS();
    let sheet = ss.getSheetByName('Messages');
    if (!sheet) {
      setupSpreadsheet();
      sheet = ss.getSheetByName('Messages');
    }

    const id = Utilities.getUuid();
    const dateStr = new Date().toISOString();
    
    // Append to sheet database
    sheet.appendRow([
      id,
      data.name.trim(),
      data.email.trim(),
      data.subject.trim(),
      data.message.trim(),
      dateStr,
      'unread', // Default status
      ''        // Notes (empty by default)
    ]);

    // Send email notification to school email
    try {
      const settings = getSettings();
      const adminEmail = settings.contactEmail || CONFIG.contactEmail || '';
      const schoolName = settings.schoolName || CONFIG.SCHOOL_NAME;
      
      if (adminEmail && adminEmail !== '-' && adminEmail.indexOf('@') !== -1) {
        const mailSubject = `[ติดต่อสอบถาม - ${schoolName}] หัวข้อ: ${data.subject}`;
        const mailBody = `มีการส่งข้อความติดต่อใหม่จากหน้าเว็บไซต์โรงเรียน\n\n` +
                         `ผู้ส่ง: ${data.name}\n` +
                         `อีเมลผู้ส่ง: ${data.email}\n` +
                         `วันที่ส่ง: ${new Date().toLocaleString('th-TH')}\n` +
                         `หัวข้อเรื่อง: ${data.subject}\n\n` +
                         `ข้อความ:\n${data.message}\n\n` +
                         `---------------------------------------------\n` +
                         `กรุณาเข้าสู่ระบบหลังบ้านเพื่อจัดการข้อความนี้`;
                         
        MailApp.sendEmail(adminEmail, mailSubject, mailBody);
      }
    } catch (mailErr) {
      Logger.log('Error sending email notification: ' + mailErr.toString());
    }

    return { success: true, id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getContactMessages(token) {
  if (!validateSession(token)) return [];
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Messages');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    
    // Sort descending by date (latest first)
    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    return sanitizeForJson(rows);
  } catch (e) {
    return [];
  }
}

function updateContactMessageStatus(token, id, status, notes) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Messages');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลข้อความ' };
    
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        const r = i + 1;
        // status is column 7 (value G), notes is column 8 (value H)
        sheet.getRange(r, 7).setValue(status || 'read');
        if (notes !== undefined) {
          sheet.getRange(r, 8).setValue(notes);
        }
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลข้อความ' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteContactMessage(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Messages');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลข้อความ' };
    
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลข้อความที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== CMS: UPDATES (📢 ข่าวประกาศอัปเดต) CRUD =====
function getUpdates(publishedOnly) {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Updates');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.id);
    
    if (publishedOnly) {
      rows = rows.filter(r => r.published === true || r.published === 'TRUE');
    }
    
    // Sort descending by date (latest first)
    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    return sanitizeForJson(rows);
  } catch (e) {
    return [];
  }
}

function addUpdate(token, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Updates');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลประกาศ' };
    
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      data.title || '',
      data.content || '',
      new Date().toISOString(),
      data.published !== false
    ]);
    return { success: true, id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateUpdate(token, id, data) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Updates');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลประกาศ' };
    
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        const r = i + 1;
        const map = { title: 2, content: 3, published: 5 };
        Object.keys(data).forEach(k => {
          if (map[k] !== undefined) {
            sheet.getRange(r, map[k]).setValue(data[k]);
          }
        });
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลประกาศ' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteUpdate(token, id) {
  if (!validateSession(token)) return { success: false, message: 'Session หมดอายุ' };
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName('Updates');
    if (!sheet) return { success: false, message: 'ไม่พบตารางข้อมูลประกาศ' };
    
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (all[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลประกาศ' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===== GEMINI API ASSISTANT SERVICES =====
function generateAiText(token, prompt) {
  if (!validateSession(token)) throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  
  try {
    const settings = getSettings();
    const provider = (settings.aiTextProvider || 'gemini').trim().toLowerCase();
    
    if (provider === 'groq') {
      const groqApiKey = (settings.groqApiKey || '').trim();
      if (!groqApiKey) {
        throw new Error('กรุณากรอกและตั้งค่า Groq API Key ในเมนู "ตั้งค่าโรงเรียน" ก่อนเริ่มใช้งาน');
      }
      
      const url = 'https://api.groq.com/openai/v1/chat/completions';
      const payload = {
        model: 'llama-3.3-70b-versatile',
        messages: [{
          role: 'user',
          content: prompt
        }],
        temperature: 0.7
      };
      
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + groqApiKey
        },
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const resText = response.getContentText();
      
      if (responseCode !== 200) {
        let errMsg = 'Groq API HTTP Error ' + responseCode;
        if (resText) {
          try {
            const resJson = JSON.parse(resText);
            if (resJson.error && resJson.error.message) {
              errMsg += ': ' + resJson.error.message;
            } else {
              errMsg += ': ' + resText.substring(0, 300);
            }
          } catch (e) {
            errMsg += ': ' + resText.substring(0, 300);
          }
        }
        throw new Error(errMsg);
      }
      
      if (!resText) {
        throw new Error('ไม่ได้รับการตอบกลับใดๆ จาก Groq API (การส่งค่าตอบกลับว่างเปล่า)');
      }
      
      let resJson;
      try {
        resJson = JSON.parse(resText);
      } catch (e) {
        throw new Error('ไม่สามารถแปลงผลลัพธ์จาก Groq เป็น JSON ได้: ' + resText.substring(0, 300));
      }
      
      if (resJson.choices && resJson.choices[0] && resJson.choices[0].message && resJson.choices[0].message.content) {
        return resJson.choices[0].message.content;
      } else {
        throw new Error('ไม่พบข้อมูลผลลัพธ์ข้อความ (choices) จาก Groq API');
      }
    } else {
      // Default to Gemini API
      const apiKey = (settings.geminiApiKey || '').trim();
      if (!apiKey) {
        throw new Error('กรุณากรอกและตั้งค่า Gemini API Key ในเมนู "ตั้งค่าโรงเรียน" ก่อนเริ่มใช้งาน');
      }
      
      const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
      const payload = {
        contents: [{
          parts: [{
            text: prompt
          }]
        }]
      };
      
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const resText = response.getContentText();
      
      if (responseCode !== 200) {
        let errMsg = 'Gemini API HTTP Error ' + responseCode;
        if (resText) {
          try {
            const resJson = JSON.parse(resText);
            if (resJson.error && resJson.error.message) {
              errMsg += ': ' + resJson.error.message;
            } else {
              errMsg += ': ' + resText.substring(0, 300);
            }
          } catch (e) {
            errMsg += ': ' + resText.substring(0, 300);
          }
        }
        throw new Error(errMsg);
      }
      
      if (!resText) {
        throw new Error('ไม่ได้รับการตอบกลับใดๆ จาก Gemini API (การส่งค่าตอบกลับว่างเปล่า)');
      }
      
      let resJson;
      try {
        resJson = JSON.parse(resText);
      } catch (e) {
        throw new Error('ไม่สามารถแปลงผลลัพธ์เป็น JSON ได้: ' + resText.substring(0, 300));
      }
      
      if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content && resJson.candidates[0].content.parts && resJson.candidates[0].content.parts[0]) {
        return resJson.candidates[0].content.parts[0].text;
      } else {
        throw new Error('ไม่พบข้อมูลผลลัพธ์ในลักษณะ Content จาก Gemini API');
      }
    }
  } catch (e) {
    throw new Error('AI Generation Error: ' + e.message);
  }
}

function generateAiImage(token, prompt) {
  if (!validateSession(token)) throw new Error('Session หมดอายุ กรุณาเข้าสู่ระบบใหม่');
  
  // Ensure the prompt guides the model to draw landscape/horizontal view
  let finalPrompt = prompt;
  if (!/landscape|horizontal|แนวนอน|16:9/i.test(prompt)) {
    finalPrompt += ', landscape orientation, widescreen 16:9';
  }
  
  let base64Bytes = '';
  let mimeType = 'image/jpeg';
  
  try {
    // Encode prompt for URL
    const encodedPrompt = encodeURIComponent(finalPrompt);
    const url = 'https://image.pollinations.ai/prompt/' + encodedPrompt + '?width=1024&height=576&nologo=true&seed=' + Math.floor(Math.random() * 1000000);
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      base64Bytes = Utilities.base64Encode(blob.getBytes());
      mimeType = blob.getContentType() || 'image/jpeg';
    } else {
      throw new Error('เซิร์ฟเวอร์ตอบกลับด้วยรหัส ' + response.getResponseCode());
    }
  } catch (e) {
    throw new Error('ไม่สามารถสร้างรูปภาพได้: ' + e.message);
  }
  
  if (base64Bytes) {
    const fileExt = mimeType.split('/')[1] || 'jpg';
    const fileName = 'ai_generated_image_' + Date.now() + '.' + fileExt;
    
    // Upload using our existing uploadImage helper
    const uploadRes = uploadImage(token, base64Bytes, fileName, mimeType);
    if (uploadRes.success) {
      return uploadRes.url; // Returns the public thumbnail url
    } else {
      throw new Error('บันทึกรูปภาพลง Google Drive ล้มเหลว: ' + uploadRes.message);
    }
  } else {
    throw new Error('ไม่พบข้อมูลรูปภาพที่สร้างขึ้นจากระบบ');
  }
}

function testImageGen() {
  try {
    const testPrompt = 'A beautiful modern Thai school campus, green lawn, clear blue sky, students smiling, 3D style';
    Logger.log('=== เริ่มต้นทดสอบสร้างรูปภาพผ่าน Pollinations.ai (ไม่มีโควตาจำกัด) ===');
    Logger.log('Prompt: "' + testPrompt + '"');
    
    // Ensure landscape orientation
    let finalPrompt = testPrompt;
    if (!/landscape|horizontal|แนวนอน|16:9/i.test(testPrompt)) {
      finalPrompt += ', landscape orientation, widescreen 16:9';
    }
    
    const encodedPrompt = encodeURIComponent(finalPrompt);
    const url = 'https://image.pollinations.ai/prompt/' + encodedPrompt + '?width=1024&height=576&nologo=true&seed=' + Math.floor(Math.random() * 1000000);
    Logger.log('เรียกใช้ URL: ' + url);
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });
    
    Logger.log('ผลลัพธ์ HTTP Code: ' + response.getResponseCode());
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      const bytes = blob.getBytes();
      const base64Bytes = Utilities.base64Encode(bytes);
      Logger.log('✅ โหลดรูปภาพสำเร็จ! ขนาดไฟล์ Base64: ' + base64Bytes.length + ' ตัวอักษร');
      Logger.log('MimeType: ' + blob.getContentType());
      Logger.log('🎉 ระบบสร้างภาพฟรีและเสถียรพร้อมใช้งานแล้ว!');
    } else {
      Logger.log('❌ โหลดรูปภาพล้มเหลว: รหัสตอบกลับ ' + response.getResponseCode());
    }
  } catch (e) {
    Logger.log('❌ เกิดข้อผิดพลาด: ' + e.message);
  }
}
