/**
 * 🛠️ Utilities: Helper Functions
 * Version: 1.3 Final
 * รวบรวมฟังก์ชันช่วยเหลือทั้งหมดจาก (V1) Utils_Common
 * ครอบคลุม: MD5, Geo Math, Text Normalization, UUID
 */

// ----------------------------------------------------
// 1. Hashing & ID Generation
// ----------------------------------------------------

/**
 * สร้าง MD5 Hash จากข้อความ (ใช้สำหรับ Cache Key ใน Service_GeoAddr)
 */
const md5 = function(key) {
  var code = key.toString().toLowerCase().replace(/\s/g, "");
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, code)
    .map(function(char) { return (char + 256).toString(16).slice(-2); })
    .join("");
};

/**
 * สร้าง UUID ใหม่ (v4)
 */
function generateUUID() {
  return Utilities.getUuid();
}

// ----------------------------------------------------
// 2. Text Processing & Normalization
// ----------------------------------------------------

/**
 * ทำความสะอาดชื่อเพื่อการเปรียบเทียบ (ตัดคำนำหน้า/สัญลักษณ์)
 */
function normalizeText(text) {
  if (!text) return "";
  var clean = text.toString().toLowerCase();
  
  // รายการคำที่ต้องการตัดออก (Stop Words) เพื่อให้เหลือแต่แก่นของชื่อ
  var stopWords = [
    "บริษัท", "บจก", "บมจ", "หจก", "ร้าน", "ห้าง", "จำกัด", 
    "มหาชน", "ส่วนบุคคล", "สาขา", "สำนักงานใหญ่", 
    "store", "shop", "company", "co.", "ltd.", 
    "จังหวัด", "อำเภอ", "ตำบล", "เขต", "แขวง", "ถนน", "ซอย", 
    "นาย", "นาง", "นางสาว", "คุณ"
  ];
  
  stopWords.forEach(function(word) {
    var regex = new RegExp(word, "g");
    clean = clean.replace(regex, "");
  });
  
  // เหลือเฉพาะตัวอักษรและตัวเลข (ลบช่องว่างและอักขระพิเศษ)
  return clean.replace(/[^a-z0-9\u0E00-\u0E7F]/g, "");
}

/**
 * ทำความสะอาดค่าระยะทางให้เป็นตัวเลขทศนิยม 2 ตำแหน่ง
 */
function cleanDistance(val) {
  if (!val && val !== 0) return "";
  var str = val.toString().replace(/[^0-9.]/g, ""); 
  var num = parseFloat(str);
  return isNaN(num) ? "" : num.toFixed(2);
}

// ----------------------------------------------------
// 3. Logic & Calculation Helpers
// ----------------------------------------------------

/**
 * เลือกชื่อที่ดีที่สุดจากกลุ่ม (Voting)
 * ใช้ใน Service_Master -> processClustering
 */
function getBestName_Smart(names) {
  var counts = {}, max = 0;
  // ✅ Fix: เริ่มต้นด้วยชื่อแรกเสมอ เพื่อกันกรณีไม่มีชื่อซ้ำเลย
  var best = (names && names.length > 0) ? names[0] : ""; 
  
  names.forEach(function(n) {
    if(!n) return;
    var k = normalizeText(n);
    counts[k] = (counts[k] || 0) + 1;
    if (counts[k] > max) { max = counts[k]; best = n; }
  });
  return best;
}

/**
 * คำนวณระยะห่างระหว่างพิกัด 2 จุด (Haversine Formula)
 * หน่วย: กิโลเมตร
 */
function getHaversineDistanceKM(lat1, lon1, lat2, lon2) {
  var R = 6371; // รัศมีโลก (กม.)
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLon/2) * Math.sin(dLon/2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

/**
 * คำนวณความเหมือนของสตริง (0.0 - 1.0)
 * ใช้ Edit Distance ในการคำนวณ
 */
function calculateSimilarity(s1, s2) {
  var longer = s1;
  var shorter = s2;
  if (s1.length < s2.length) {
    longer = s2;
    shorter = s1;
  }
  var longerLength = longer.length;
  if (longerLength === 0) {
    return 1.0;
  }
  return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}

/**
 * Levenshtein Edit Distance Algorithm
 * ใช้ช่วยคำนวณความต่างของคำ
 */
function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  var costs = new Array();
  for (var i = 0; i <= s1.length; i++) {
    var lastValue = i;
    for (var j = 0; j <= s2.length; j++) {
      if (i == 0)
        costs[j] = j;
      else {
        if (j > 0) {
          var newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

