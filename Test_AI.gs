function forceRunAI_Now() {
  // บังคับ AutoPilot ให้ทำงานทันที โดยไม่สน Limit
  processAIIndexing(); 
  SpreadsheetApp.getUi().alert("✅ สั่ง AI วิเคราะห์แล้ว!\nลองกลับไปดูที่ชีต Database คอลัมน์ F ว่ามีคำว่า 'บีขวิก' หรือยัง");
}
