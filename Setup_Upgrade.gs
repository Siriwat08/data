/**
 * 🛠️ System Upgrade Tool
 * หน้าที่: อัปเกรดโครงสร้างชีต Database จาก V1 -> V2
 * (เพิ่มคอลัมน์ใหม่ที่จำเป็นสำหรับระบบอัจฉริยะ)
 */

function upgradeDatabaseStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME); // "Database"
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ ไม่พบชีต Database");
    return;
  }

  // รายชื่อคอลัมน์ใหม่ที่จะเพิ่ม (ต่อท้ายจากเดิม)
  var newHeaders = [
    "Customer Type",      // Col 16: ประเภทลูกค้า (VIP, B2B)
    "Time Window",        // Col 17: เวลารับของ (08:00-17:00)
    "Avg Service Time",   // Col 18: เวลาลงของเฉลี่ย (นาที)
    "Vehicle Constraint", // Col 19: ข้อจำกัดรถ (4W Only)
    "Contact Person",     // Col 20: ชื่อผู้ติดต่อ
    "Phone Number",       // Col 21: เบอร์โทร
    "Risk Score",         // Col 22: ความเสี่ยง (0-10)
    "Branch Code",        // Col 23: รหัสสาขา
    "Last Updated"        // Col 24: อัปเดตล่าสุดเมื่อ
  ];

  var lastCol = sheet.getLastColumn();
  
  // เช็คว่าอัปเกรดไปหรือยัง (ถ้ามีเกิน 15 คอลัมน์แสดงว่าอาจจะเคยทำแล้ว)
  if (lastCol > 15) {
    var response = SpreadsheetApp.getUi().alert(
      "⚠️ ตรวจสอบ", 
      "ดูเหมือนชีต Database จะมีคอลัมน์มากกว่า 15 แล้ว ต้องการเพิ่มต่อท้ายอีกหรือไม่?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (response == SpreadsheetApp.getUi().Button.NO) return;
  }

  // เริ่มสร้างหัวตารางใหม่
  var startCol = lastCol + 1;
  var range = sheet.getRange(1, startCol, 1, newHeaders.length);
  
  range.setValues([newHeaders]);
  range.setFontWeight("bold");
  range.setBackground("#e6f7ff"); // สีพื้นหลังให้รู้ว่าเป็นของใหม่
  
  // จัด Format
  sheet.autoResizeColumns(startCol, newHeaders.length);
  
  SpreadsheetApp.getUi().alert("✅ อัปเกรดฐานข้อมูลเรียบร้อย!\nเพิ่มคอลัมน์ใหม่สำหรับรองรับระบบ V2 แล้ว");
}

/**
 * 🔍 ฟังก์ชันตรวจสอบข้อมูลซ้ำ (Smart Deduplicate)
 * เอาไว้เช็คว่า ชื่อต่างกัน แต่พิกัดเดียวกันหรือไม่
 */
function findHiddenDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var duplicates = [];
  
  // Loop ตรวจสอบ (ข้าม Header)
  for (var i = 1; i < data.length; i++) {
    for (var j = i + 1; j < data.length; j++) {
      var row1 = data[i];
      var row2 = data[j];
      
      // ถ้าพิกัดใกล้กันมาก (ระยะห่าง < 50 เมตร) แต่ชื่อไม่เหมือนกัน
      var dist = getHaversineDistanceKM(row1[1], row1[2], row2[1], row2[2]); // ใช้ฟังก์ชันเดิมที่มีใน V1
      
      if (dist < 0.05) { // 50 เมตร
        duplicates.push({
          row1: i + 1,
          name1: row1[0],
          row2: j + 1,
          name2: row2[0],
          distance: (dist * 1000).toFixed(0) + " เมตร"
        });
      }
    }
  }
  
  if (duplicates.length > 0) {
    var msg = "⚠️ พบข้อมูลที่น่าจะซ้ำกัน " + duplicates.length + " คู่:\n";
    duplicates.slice(0, 10).forEach(d => { // โชว์แค่ 10 อันแรก
      msg += `- ${d.name1} vs ${d.name2} (ห่าง ${d.distance})\n`;
    });
    SpreadsheetApp.getUi().alert(msg);
  } else {
    SpreadsheetApp.getUi().alert("✅ ไม่พบข้อมูลซ้ำซ้อนในระยะใกล้");
  }
}
