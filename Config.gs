/**
 * ⚙️ รายชื่อตัวแปรและค่าคงที่ (Global Config)
 * อ้างอิงจากเอกสาร (V1) Logistics Master Data System
 * เพิ่มเติม: รองรับ Col P (Created), Q (Updated) และ AI Config
 */

var CONFIG = {
  SHEET_NAME: "Database",
  MAPPING_SHEET: "NameMapping",
  SOURCE_SHEET: "SCGนครหลวงJWDภูมิภาค",
  
  // 🧠 AI CONFIGURATION
  // ไปขอ Key ฟรีได้ที่: https://aistudio.google.com/app/apikey
  GEMINI_API_KEY: "xxx", 
  USE_AI_AUTO_FIX: true, // เปิดให้ AI ช่วยแก้ที่อยู่หรือไม่

  // 🔴 พิกัดคลังสินค้า (Center Point)
  DEPOT_LAT: 14.164688, 
  DEPOT_LNG: 100.625354,

  // คอลัมน์ Master (Index เริ่มที่ 1 = A)
  COL_NAME: 1,      // A: ชื่อลูกค้า
  COL_LAT: 2,       // B: Latitude
  COL_LNG: 3,       // C: Longitude
  COL_SUGGESTED: 4, // D: ชื่อที่ระบบแนะนำ
  COL_CONFIDENCE: 5,// E: ความมั่นใจ
  COL_NORMALIZED: 6,// F: ชื่อที่ Clean แล้ว
  COL_VERIFIED: 7,  // G: สถานะตรวจสอบ (Checkbox)
  COL_SYS_ADDR: 8,  // H: ที่อยู่จากระบบต้นทาง
  COL_ADDR_GOOG: 9, // I: ที่อยู่จาก Google Maps
  COL_DIST_KM: 10,  // J: ระยะทางจากคลัง
  
  // Enterprise Columns (UUID & Meta)
  COL_UUID: 11,     // K: Unique ID
  COL_PROVINCE: 12, // L: จังหวัด
  COL_DISTRICT: 13, // M: อำเภอ
  COL_POSTCODE: 14, // N: รหัสไปรษณีย์
  COL_QUALITY: 15,  // O: Quality Score
  COL_CREATED: 16,  // P: วันที่สร้าง (Created)
  COL_UPDATED: 17,  // Q: วันที่แก้ไขล่าสุด (Updated)

  DISTANCE_THRESHOLD_KM: 0.05, 
  BATCH_LIMIT: 50,  
  DEEP_CLEAN_LIMIT: 100 
};

// Config สำหรับ SCG API & Daily Operation
const SCG_CONFIG = {
  SHEET_DATA: 'Data',
  SHEET_INPUT: 'Input',
  SHEET_EMPLOYEE: 'ข้อมูลพนักงาน',
  API_URL: 'https://fsm.scgjwd.com/Monitor/SearchDelivery',
  INPUT_START_ROW: 4,
  COOKIE_CELL: 'B1',
  SHIPMENT_STRING_CELL: 'B3',
  SHEET_MASTER_DB: 'Database',
  SHEET_MAPPING: 'NameMapping'
};
