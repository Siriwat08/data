/**
 * 🕵️ Service: Logistics AI Agent (Final Integrated)
 * Codename: "The Steward"
 * หน้าที่: เป็น Agent อัจฉริยะ คอยคาดการณ์คำค้นหาและการพิมพ์ผิด (Typos) ล่วงหน้า
 */

var AGENT_CONFIG = {
  NAME: "Logistics_Agent_01",
  MODEL: "gemini-1.5-flash",
  BATCH_SIZE: 3, // ทำทีละ 3 เจ้า (เพื่อความเสถียร)
  TAG: "[Agent_Ver2]" // เอาไว้แปะป้ายว่าตรวจแล้ว
};

/**
 * 👋 สั่ง Agent ให้ตื่นมาทำงานเดี๋ยวนี้ (Manual Trigger)
 */
function WAKE_UP_AGENT() {
  SpreadsheetApp.getUi().toast("🕵️ Agent: ผมตื่นแล้วครับ กำลังเริ่มวิเคราะห์ข้อมูล...", "AI Agent Started");
  runAgentLoop();
  SpreadsheetApp.getUi().alert("✅ Agent รายงานผล:\nผมวิเคราะห์ข้อมูลชุดล่าสุดเสร็จแล้วครับ ลองไปค้นหาดูได้เลย!");
}

/**
 * ⏰ ตั้งเวลาให้ Agent ตื่นมาทำงานเองทุก 10 นาที
 */
function SCHEDULE_AGENT_WORK() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runAgentLoop") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger("runAgentLoop")
    .timeBased()
    .everyMinutes(10)
    .create();
    
  SpreadsheetApp.getUi().alert("✅ ตั้งค่าเรียบร้อย!\nAgent จะตื่นมาทำงานทุก 10 นาที เพื่อเตรียมข้อมูลให้ท่านครับ");
}

/**
 * 🔄 Agent Loop (กระบวนการคิดของ AI)
 */
function runAgentLoop() {
  console.time("Agent_Thinking_Time");
  
  try {
    if (!CONFIG.GEMINI_API_KEY) {
      console.error("Agent: เจ้านายครับ ผมไม่มีกุญแจ (API Key) ผมเข้า Gemini ไม่ได้ครับ");
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME); // Database
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    // อ่านข้อมูลมาวิเคราะห์ (Col A ถึง Col O/P)
    // หมายเหตุ: ปรับช่วง Column ตาม Config จริงของท่าน
    var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()); 
    var data = dataRange.getValues();
    var jobsDone = 0;

    for (var i = 0; i < data.length; i++) {
      if (jobsDone >= AGENT_CONFIG.BATCH_SIZE) break;

      var row = data[i];
      var name = row[CONFIG.COL_NAME - 1];
      var currentNorm = row[CONFIG.COL_NORMALIZED - 1]; // ช่องที่ Agent จะเขียน (Col F)
      
      // เงื่อนไข: ถ้ามีชื่อ แต่ยังไม่มีลายเซ็น Agent หรือข้อมูลว่าง
      if (name && (!currentNorm || String(currentNorm).indexOf(AGENT_CONFIG.TAG) === -1)) {
        
        console.log(`Agent: กำลังเพ่งเล็งเป้าหมาย "${name}"...`);
        
        // 🧠 ใช้สมอง AI คิดวิเคราะห์คำผิด/คำค้นหา
        var aiThoughts = askGeminiToPredictTypos(name);
        
        // 📝 บันทึกผลลัพธ์ลง Database
        var knowledgeBase = name + " " + aiThoughts + " " + AGENT_CONFIG.TAG;
        sheet.getRange(i + 2, CONFIG.COL_NORMALIZED).setValue(knowledgeBase);
        
        // 🆔 แถม: เติม UUID ให้ด้วยถ้าไม่มี
        var uuidIdx = (CONFIG.COL_UUID || 15) - 1;
        if (!row[uuidIdx]) {
          sheet.getRange(i + 2, CONFIG.COL_UUID).setValue(Utilities.getUuid());
        }

        console.log(`Agent: ✅ เรียนรู้สำเร็จ! คาดเดาคำว่า -> ${aiThoughts}`);
        jobsDone++;
      }
    }
    
    // ทำงาน Sync งานเดิมไปด้วย (ถ้ามีฟังก์ชันนี้)
    if (typeof applyMasterCoordinatesToDailyJob === 'function') {
       applyMasterCoordinatesToDailyJob();
    }

  } catch (e) {
    console.error("Agent: เกิดข้อผิดพลาด! " + e.message);
  }
  
  console.timeEnd("Agent_Thinking_Time");
}

/**
 * 📡 Skill: การคาดเดาคำผิด (Typos Prediction)
 */
function askGeminiToPredictTypos(originalName) {
  var prompt = `
    Task: You are a Thai Logistics Search Agent.
    Input Name: "${originalName}"
    Goal: Generate a list of search keywords including common typos, phonetic spellings, and abbreviations.
    Constraint: Output ONLY the keywords separated by spaces.
    Example Input: "บี-ควิก (สาขาลาดพร้าว)"
    Example Output: บีควิก บีขวิก บีวิก BeQuik BQuik B-Quik ลาดพร้าว BQuick
  `;

  try {
    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "temperature": 0.4 }
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var url = `https://generativelanguage.googleapis.com/v1beta/models/${AGENT_CONFIG.MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates[0].content) {
      return json.candidates[0].content.parts[0].text.trim();
    }
  } catch (e) {
    console.warn("Agent Error: " + e.message);
  }
  
  // Fallback Logic ถ้า AI ป่วย
  return (typeof normalizeText === 'function') ? normalizeText(originalName) : originalName;
}
