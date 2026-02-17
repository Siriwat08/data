/**
 * 🤖 Service: Auto Pilot (AI Edition)
 * Version: 3.0 Gemini Integration
 * หน้าที่: คู่หูอัจฉริยะ ทำงานเบื้องหลังด้วย Google Gemini AI
 * * 1. Sync ข้อมูล SCG (งานประจำ)
 * 2. 🧠 AI Smart Indexing: ใช้ Gemini วิเคราะห์ชื่อร้าน หาคำพ้อง/ชื่อย่อ/คำค้นหาที่น่าจะเป็น
 * 3. Auto-Fix: เติมพิกัดที่ขาด
 */

/**
 * ▶️ ฟังก์ชันเปิดระบบ Auto-Pilot
 */
function START_AUTO_PILOT() {
  STOP_AUTO_PILOT();
  
  ScriptApp.newTrigger("autoPilotRoutine")
    .timeBased()
    .everyMinutes(10) // ทำงานทุก 10 นาที
    .create();
    
  SpreadsheetApp.getUi().alert("▶️ AI Auto-Pilot: เปิดระบบแล้ว\n(ผมจะใช้ Gemini ช่วยวิเคราะห์ข้อมูลทุกๆ 10 นาทีครับ)");
}

function STOP_AUTO_PILOT() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "autoPilotRoutine") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 🔄 Main Routine
 */
function autoPilotRoutine() {
  // 1. งาน SCG (คงเดิม)
  try {
    if (typeof applyMasterCoordinatesToDailyJob === 'function') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var dataSheet = ss.getSheetByName(CONFIG.SHEET_DATA || "Data");
      if (dataSheet && dataSheet.getLastRow() > 1) {
         applyMasterCoordinatesToDailyJob();
         console.log("AutoPilot: SCG Sync Done.");
      }
    }
  } catch(e) { console.error("SCG Error: " + e.message); }

  // 2. งาน AI (พระเอกของเรา)
  try {
    processAIIndexing();
    console.log("AutoPilot: AI Indexing Done.");
  } catch(e) { console.error("AI Error: " + e.message); }
}

/**
 * 🧠 AI Processing Logic
 * ดึงข้อมูลมาให้ Gemini ช่วยคิดคำค้นหา (Keywords)
 */
function processAIIndexing() {
  // ตรวจสอบ Key ก่อน
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY.length < 10) {
    console.log("⚠️ ข้ามการทำงาน AI เพราะยังไม่ได้ใส่ GEMINI_API_KEY ใน Config");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // อ่านข้อมูล (เฉพาะ Col ชื่อ และ Col Normalized)
  // สมมติ COL_NAME=1, COL_NORMALIZED=6
  var rangeName = sheet.getRange(2, CONFIG.COL_NAME, lastRow - 1, 1);
  var rangeNorm = sheet.getRange(2, CONFIG.COL_NORMALIZED, lastRow - 1, 1);
  
  var names = rangeName.getValues();
  var norms = rangeNorm.getValues();
  
  var aiCount = 0;
  var AI_LIMIT = 3; // ⚠️ ทำทีละ 3 เจ้าพอ (กัน Quota เต็ม/ระบบค้าง)

  for (var i = 0; i < names.length; i++) {
    if (aiCount >= AI_LIMIT) break;

    var name = names[i][0];
    var currentNorm = norms[i][0];

    // เงื่อนไข: มีชื่อ แต่ยังไม่มี Tag "[AI]" ในช่อง Normalized
    // หรือข้อมูลว่างเปล่า
    if (name && (!currentNorm || currentNorm.toString().indexOf("[AI]") === -1)) {
      
      // 1. สร้าง Basic Index ก่อน (กันเหนียว)
      var basicKey = createBasicSmartKey(name);
      
      // 2. เรียก Gemini ให้ช่วยคิด (นี่คือ AI จริงๆ)
      var aiKeywords = callGeminiThinking(name);
      
      // 3. รวมร่าง: Basic + AI Keywords
      // ใส่ Tag [AI] ไว้ท้ายสุด เพื่อบอกว่าแถวนี้ AI ตรวจแล้ว รอบหน้าจะได้ไม่ทำซ้ำ
      var finalString = basicKey + " " + aiKeywords + " [AI]";
      
      // อัปเดต Array (และบันทึกลง Sheet ทันทีเพื่อกันพลาด)
      sheet.getRange(i + 2, CONFIG.COL_NORMALIZED).setValue(finalString);
      
      console.log(`🤖 AI Analyzed: ${name} -> ${aiKeywords}`);
      aiCount++;
    }
  }
}

/**
 * 📡 ฟังก์ชันเรียก Gemini API
 */
function callGeminiThinking(customerName) {
  try {
    var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + CONFIG.GEMINI_API_KEY;
    
    // Prompt สั่งงาน AI
    var prompt = `
      คุณคือผู้ช่วย Logistics อัจฉริยะ
      วิเคราะห์ชื่อลูกค้า: "${customerName}"
      
      หน้าที่ของคุณ: 
      1. เดา "คำค้นหา" (Keywords) ที่คนขับรถอาจจะใช้ค้นหาที่นี่
      2. ถ้าเป็นชื่อย่อ ให้ขยายความ (เช่น รพ. -> โรงพยาบาล)
      3. ถ้าเป็นภาษาอังกฤษ ให้ขอคำอ่านไทย หรือถ้าเป็นไทย ให้ขอคำทับศัพท์
      4. ขอสั้นๆ ไม่เกิน 5 คำ คั่นด้วยเว้นวรรค
      
      ตัวอย่าง:
      Input: "บจก. เอสซีจี (สาขาบางซื่อ)"
      Output: SCG ปูนใหญ่ บางซื่อ SiamCement
      
      Output ของคุณ (เฉพาะคำค้นหา ไม่เอาคำอธิบาย):
    `;

    var payload = {
      "contents": [{
        "parts": [{ "text": prompt }]
      }]
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    var json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates.length > 0) {
      var text = json.candidates[0].content.parts[0].text;
      // ล้าง Format ที่ AI อาจจะแถมมา (เช่น \n หรือ *)
      return text.replace(/\n/g, " ").replace(/\*/g, "").trim();
    }
  } catch (e) {
    console.warn("Gemini Error: " + e.message);
    return ""; // ถ้า AI ป่วย ให้คืนค่าว่างไปก่อน ไม่ให้ระบบล่ม
  }
  return "";
}

/**
 * 🔨 Helper: สร้าง Index แบบพื้นฐาน (Regex)
 * เอาไว้กันเหนียว ช่วงที่รอ AI ทำงาน
 */
function createBasicSmartKey(text) {
  if (!text) return "";
  // ลบ บจก., ช่องว่าง, อักขระพิเศษ
  var clean = text.toString().replace(/\s+/g, '').replace(/^(บจก|หจก|ร้าน|บริษัท)\.?/g, '');
  // ลบวรรณยุกต์ (Anti-Typo)
  var noTones = clean.replace(/[\u0E48-\u0E4C]/g, "");
  
  if (clean === noTones) return clean;
  return clean + " " + noTones;
}
