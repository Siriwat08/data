/**
 * 🌐 WebApp Controller
 * หน้าที่: แสดงผลหน้าเว็บ (doGet) และรวมไฟล์ HTML
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('🔍 Logistics Search Engine')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับดึง CSS/JS เข้ามาใน HTML (ถ้าแยกไฟล์)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
