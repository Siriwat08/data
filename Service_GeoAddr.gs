/**
 * 🌍 Service: Geo Address
 * Version: 1.3 Final (Restored User's Custom Logic)
 * Features:
 * 1. PostalRef Caching: อ่านข้อมูลรหัสไปรษณีย์จากชีต PostalRef
 * 2. Geocoding: แปลงพิกัดเป็นที่อยู่ (Reverse Geocode)
 * 3. Distance: คำนวณระยะทางขับรถจริง (Driving Mode)
 * 4. Caching: ใช้ DocumentCache + MD5
 */

var _POSTAL_CACHE = null;

function parseAddressFromText(fullAddress) {
  var result = { province: "", district: "", postcode: "" };
  if (!fullAddress) return result;
  
  var postalDB = getPostalDataCached();
  if (!postalDB) return result;
  
  // ค้นหารหัสไปรษณีย์ 5 หลัก
  var zipMatch = fullAddress.toString().match(/(\d{5})/);
  if (zipMatch && postalDB.byZip[zipMatch[5]]) {
    var infoList = postalDB.byZip[zipMatch[5]];
    if (infoList.length > 0) {
       // ใช้ข้อมูลตัวแรกที่เจอใน DB
       var info = infoList[0];
       return { province: info.province, district: info.district, postcode: zipMatch[5] };
    }
  }
  return result;
}

function getPostalDataCached() {
  if (_POSTAL_CACHE) return _POSTAL_CACHE;
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PostalRef");
  if (!sheet) return null;
  
  // อ่านข้อมูลทั้งหมด (ใช้ DataRange เพื่อให้แน่ใจว่าอ่านครบทุกคอลัมน์ถึง I)
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var db = { byZip: {} };
  
  data.forEach(row => {
    // สมมติ: Col A (0) = Postcode
    var pc = String(row[0]).trim(); 
    if (!db.byZip[pc]) db.byZip[pc] = [];
    
    // อ้างอิง Index เดิมของท่าน: row[6]=District, row[8]=Province
    // (ตรวจสอบว่าแถวนี้มีข้อมูลถึง Col I ไหม ถ้าไม่มีให้ข้าม)
    if (row.length > 8) {
      db.byZip[pc].push({ postcode: pc, district: row[6], province: row[8] });
    }
  });
  
  _POSTAL_CACHE = db;
  return db;
}

// ---------------------------
// MAPS API & CACHE
// ---------------------------

function GET_ADDR_WITH_CACHE(lat, lng) {
  if (!lat || !lng) return "";
  var key = "rev_" + lat + "_" + lng;
  var cached = getCache(key);
  if (cached) return cached;

  try {
    var response = Maps.newGeocoder().setLanguage("th").reverseGeocode(lat, lng);
    if (response.results && response.results.length > 0) {
      var addr = response.results[0].formatted_address;
      setCache(key, addr);
      return addr;
    }
  } catch (e) {}
  return "";
}

function CALCULATE_DISTANCE_KM(origin, destination) {
  if (!origin || !destination) return "";
  var key = "dist_" + origin + "_" + destination;
  var cached = getCache(key);
  if (cached) return cached;

  try {
    // ใช้ DirectionFinder (Driving Mode) ตามโค้ดเดิมของท่าน
    var directions = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(destination)
      .setMode(Maps.DirectionFinder.Mode.DRIVING) 
      .getDirections();

    if (directions.routes && directions.routes.length > 0) {
      var legs = directions.routes[0].legs;
      if (legs && legs.length > 0) {
        var meters = legs[0].distance.value;
        var km = (meters / 1000).toFixed(2); 
        setCache(key, km);
        return km;
      }
    }
  } catch (e) {}
  return "";
}

// Cache Helper using DocumentCache & MD5
// หมายเหตุ: ฟังก์ชัน md5() ต้องมีอยู่ในไฟล์ Utils_Common.gs
const getCache = key => CacheService.getDocumentCache().get(md5(key));
const setCache = (key, value) => CacheService.getDocumentCache().put(md5(key), value, 21600);
