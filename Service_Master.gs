/**
 * 🧠 Service: Master Data Management
 * Version: 1.3 Final (Full Logic Restored)
 * Features:
 * 1. Sync: ดึงข้อมูลใหม่ + สร้าง UUID + Timestamp
 * 2. Deep Clean: ตรวจสอบพิกัด, ที่อยู่, ระยะทาง, แยกจังหวัด/อำเภอ
 * 3. Finalize: จบงาน ย้ายไป Mapping พร้อม UUID
 * 4. Clustering: จัดกลุ่มพิกัดใกล้เคียง (Logic เต็ม)
 * 5. Admin Tools: ซ่อม UUID และ Mapping
 */

// ==========================================
// 1. IMPORT & SYNC
// ==========================================

function syncNewDataToMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sourceSheet || !masterSheet) { Browser.msgBox("❌ ไม่พบ Sheet (Source หรือ Database)"); return; }

  // Mapping Column จาก Source (SCGนครหลวง...)
  var SRC = { NAME: 13, LAT: 15, LNG: 16, SYS_ADDR: 19, DIST: 24, GOOG_ADDR: 25 };

  var lastRowM = masterSheet.getLastRow();
  var existingNames = {};
  
  // โหลดชื่อเดิมเพื่อกันซ้ำ
  if (lastRowM > 1) {
    var mData = masterSheet.getRange(2, CONFIG.COL_NAME, lastRowM - 1, 1).getValues();
    mData.forEach(function(r) { if (r[0]) existingNames[normalizeText(r[0])] = true; });
  }

  var lastRowS = sourceSheet.getLastRow();
  if (lastRowS < 2) return;
  
  var sData = sourceSheet.getRange(2, 1, lastRowS - 1, 25).getValues();
  var newEntries = [];
  var currentBatch = {};

  sData.forEach(function(row) {
    var name = row[SRC.NAME - 1];
    var lat = row[SRC.LAT - 1];
    var lng = row[SRC.LNG - 1];
    
    if (!name || !lat || !lng) return;
    
    var clean = normalizeText(name);
    // เช็คซ้ำทั้งใน DB และใน Batch ปัจจุบัน
    if (!existingNames[clean] && !currentBatch[clean]) {
      var newRow = new Array(17).fill(""); // จองพื้นที่ถึง Col Q
      
      newRow[CONFIG.COL_NAME - 1] = name;
      newRow[CONFIG.COL_LAT - 1] = lat;
      newRow[CONFIG.COL_LNG - 1] = lng;
      newRow[CONFIG.COL_VERIFIED - 1] = false; 
      newRow[CONFIG.COL_SYS_ADDR - 1] = row[SRC.SYS_ADDR - 1]; 
      newRow[CONFIG.COL_ADDR_GOOG - 1] = row[SRC.GOOG_ADDR - 1]; 
      newRow[CONFIG.COL_DIST_KM - 1] = cleanDistance(row[SRC.DIST - 1]); 
      
      // Enterprise Data
      newRow[CONFIG.COL_UUID - 1] = generateUUID(); 
      newRow[CONFIG.COL_CREATED - 1] = new Date(); 
      newRow[CONFIG.COL_UPDATED - 1] = new Date();
      
      newEntries.push(newRow);
      currentBatch[clean] = true;
    }
  });

  if (newEntries.length > 0) {
    masterSheet.getRange(lastRowM + 1, 1, newEntries.length, 17).setValues(newEntries);
    Browser.msgBox("✅ นำเข้าข้อมูลใหม่ " + newEntries.length + " รายการ");
  } else {
    Browser.msgBox("👌 ไม่มีข้อมูลใหม่ที่ต้องนำเข้า");
  }
}

// ==========================================
// 2. DATA ENRICHMENT (GEO & CLUSTER)
// ==========================================

function updateGeoData_SmartCache() { 
  // เรียก DeepClean แบบจำกัด Scope หรือจะแยก Logic ก็ได้
  // ใน V1.3 เราใช้ DeepClean เป็นตัวหลักในการซ่อมข้อมูล
  runDeepCleanBatch_100(); 
}

function autoGenerateMasterList_Smart() { 
  processClustering(); 
}

// ==========================================
// 3. DEEP CLEAN & VALIDATION
// ==========================================

function runDeepCleanBatch_100() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var props = PropertiesService.getScriptProperties();
  var startRow = parseInt(props.getProperty('DEEP_CLEAN_POINTER') || '2');
  
  if (startRow > lastRow) {
    Browser.msgBox("🎉 ตรวจครบทุกแถวแล้วครับ! (กดรีเซ็ตถ้าต้องการเริ่มใหม่)");
    return;
  }

  var endRow = Math.min(startRow + CONFIG.DEEP_CLEAN_LIMIT - 1, lastRow);
  var numRows = endRow - startRow + 1;
  
  // อ่านข้อมูลถึง Col Q (17)
  var range = sheet.getRange(startRow, 1, numRows, 17);
  var values = range.getValues();
  
  var origin = CONFIG.DEPOT_LAT + "," + CONFIG.DEPOT_LNG;
  var updatedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var lat = row[CONFIG.COL_LAT - 1];
    var lng = row[CONFIG.COL_LNG - 1];
    var googleAddr = row[CONFIG.COL_ADDR_GOOG - 1];
    var distKM = row[CONFIG.COL_DIST_KM - 1];
    var hasCoord = (lat && lng && !isNaN(lat) && !isNaN(lng));
    var changed = false;

    // Task A: เติมที่อยู่และระยะทาง (ถ้าขาด)
    if (hasCoord) {
      if (!googleAddr || googleAddr === "") {
        var addr = GET_ADDR_WITH_CACHE(lat, lng);
        if (addr && addr !== "Error") {
          row[CONFIG.COL_ADDR_GOOG - 1] = addr;
          googleAddr = addr; // อัปเดตตัวแปร local เพื่อใช้ต่อ
          changed = true;
        }
      }
      if (!distKM || distKM === "") {
        var km = CALCULATE_DISTANCE_KM(origin, lat + "," + lng);
        if (km) { 
          row[CONFIG.COL_DIST_KM - 1] = km; 
          changed = true; 
        }
      }
    }
    
    // Task B: เติม UUID (ถ้าขาด)
    if (!row[CONFIG.COL_UUID - 1]) { 
      row[CONFIG.COL_UUID - 1] = generateUUID(); 
      row[CONFIG.COL_CREATED - 1] = row[CONFIG.COL_CREATED - 1] || new Date(); 
      changed = true; 
    }

    // Task C: แกะที่อยู่ลง Col L, M, N (โดยใช้ Service_GeoAddr)
    if (googleAddr && (!row[CONFIG.COL_PROVINCE - 1] || !row[CONFIG.COL_DISTRICT - 1])) {
       // เรียกใช้ parseAddressFromText จาก Service_GeoAddr.gs
       var parsed = parseAddressFromText(googleAddr);
       if (parsed.province) {
         row[CONFIG.COL_PROVINCE - 1] = parsed.province;
         row[CONFIG.COL_DISTRICT - 1] = parsed.district;
         row[CONFIG.COL_POSTCODE - 1] = parsed.postcode;
         changed = true;
       }
    }

    if (changed) {
       row[CONFIG.COL_UPDATED - 1] = new Date(); // Update timestamp
       updatedCount++;
    }
  }

  if (updatedCount > 0) {
    range.setValues(values);
  }
  
  props.setProperty('DEEP_CLEAN_POINTER', (endRow + 1).toString());
  ss.toast("✅ ตรวจสอบช่วงแถว " + startRow + " ถึง " + endRow + "\n(แก้ไข " + updatedCount + " รายการ)", "Deep Clean Status");
}

function resetDeepCleanMemory() {
  PropertiesService.getScriptProperties().deleteProperty('DEEP_CLEAN_POINTER');
  Browser.msgBox("🔄 รีเซ็ตความจำแล้ว เริ่มต้นใหม่ที่แถว 2");
}

// ==========================================
// 4. FINALIZE & MAPPING
// ==========================================

function finalizeAndClean_MoveToMapping() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  
  if (!masterSheet || !mapSheet) { Browser.msgBox("❌ ไม่พบ Sheet"); return; }
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  var uuidMap = {};
  var allData = masterSheet.getRange(2, 1, lastRow - 1, 17).getValues();
  
  // สร้าง Map ของ UUID เดิมที่มีอยู่
  allData.forEach(function(row) {
    var name = normalizeText(row[CONFIG.COL_NAME - 1]);
    var suggested = normalizeText(row[CONFIG.COL_SUGGESTED - 1]);
    var uuid = row[CONFIG.COL_UUID - 1];
    
    if (uuid) {
      if (name) uuidMap[name] = uuid;
      if (suggested) uuidMap[suggested] = uuid; 
    }
  });

  // Backup ข้อมูลก่อนลบ
  var backupName = "Backup_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss");
  masterSheet.copyTo(ss).setName(backupName);
  
  var rowsToKeep = [];       
  var mappingToUpload = []; 
  var processedNames = {}; 

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    var rawName = row[CONFIG.COL_NAME - 1];      
    var suggestedName = row[CONFIG.COL_SUGGESTED - 1]; 
    var isVerified = row[CONFIG.COL_VERIFIED - 1];    
    var currentUUID = row[CONFIG.COL_UUID - 1];

    if (isVerified === true) {
      // ถ้า Verified แล้ว เก็บไว้ใน Master
      rowsToKeep.push(row); 
    } 
    else if (suggestedName && suggestedName !== "") {
      // ถ้ามี Suggestion ย้ายไป Mapping
      if (rawName !== suggestedName && !processedNames[rawName]) {
        var targetUUID = uuidMap[normalizeText(suggestedName)] || currentUUID;
        mappingToUpload.push([rawName, suggestedName, targetUUID]);
        processedNames[rawName] = true;
      }
    }
    // กรณีอื่นๆ (ยังไม่ได้ Verify และไม่มี Suggest) จะถูกลบออก (แต่มี Backup แล้ว)
  }

  // บันทึก Mapping
  if (mappingToUpload.length > 0) {
    var lastRowMap = mapSheet.getLastRow();
    var existingMapKeys = {};
    if (lastRowMap > 1) {
      var mapData = mapSheet.getRange(2, 1, lastRowMap - 1, 1).getValues();
      mapData.forEach(function(r) { existingMapKeys[normalizeText(r[0])] = true; });
    }
    var finalMapping = mappingToUpload.filter(function(m) { return !existingMapKeys[normalizeText(m[0])]; });
    
    if (finalMapping.length > 0) {
      mapSheet.getRange(mapSheet.getLastRow() + 1, 1, finalMapping.length, 3).setValues(finalMapping);
    }
  }

  // เขียนข้อมูล Master ใหม่ (เฉพาะ Verified)
  masterSheet.getRange(2, 1, lastRow, 17).clearContent(); 
  
  if (rowsToKeep.length > 0) {
    masterSheet.getRange(2, 1, rowsToKeep.length, 17).setValues(rowsToKeep);
    Browser.msgBox("✅ จบงานเรียบร้อย!\n- เพิ่ม Mapping: " + mappingToUpload.length + " รายการ\n- คงเหลือ Master: " + rowsToKeep.length + " รายการ");
  } else {
    masterSheet.getRange(2, 1, allData.length, 17).setValues(allData); // กู้คืนถ้าไม่มี Verified เลย
    Browser.msgBox("⚠️ ไม่พบข้อมูล Verified เลย (ระบบได้กู้คืนข้อมูลเดิมกลับมาให้แล้ว)");
  }
}

// ==========================================
// 5. ADMIN TOOLS
// ==========================================

function assignMissingUUIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  if (sheet.getMaxColumns() < CONFIG.COL_UUID) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), CONFIG.COL_UUID - sheet.getMaxColumns());
    sheet.getRange(1, CONFIG.COL_UUID).setValue("UUID").setFontWeight("bold");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, CONFIG.COL_UUID, lastRow - 1, 1);
  var values = range.getValues();
  var count = 0;

  var newValues = values.map(function(r) {
    if (!r[0]) {
      count++;
      return [generateUUID()];
    }
    return [r[0]];
  });

  if (count > 0) {
    range.setValues(newValues);
    Browser.msgBox("✅ สร้าง UUID ใหม่จำนวน: " + count);
  } else {
    Browser.msgBox("ℹ️ ข้อมูลครบถ้วนแล้ว ไม่มีการสร้างเพิ่ม");
  }
}

function repairNameMapping_Full() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  
  if (!dbSheet || !mapSheet) { Browser.msgBox("❌ ไม่พบ Sheet"); return; }

  // 1. ดึง UUID จาก Database
  var dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, CONFIG.COL_UUID).getValues();
  var uuidMap = {};
  dbData.forEach(function(r) {
    if (r[CONFIG.COL_UUID-1]) {
       uuidMap[normalizeText(r[CONFIG.COL_NAME-1])] = r[CONFIG.COL_UUID-1];
    }
  });

  // 2. ตรวจสอบ Mapping
  var mapRange = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 3);
  var mapValues = mapRange.getValues();
  var cleanList = [];
  var seen = {};

  mapValues.forEach(function(r) {
    var oldN = r[0], newN = r[1], uid = r[2];
    var normOld = normalizeText(oldN);
    
    if (!normOld) return; // ข้ามแถวว่าง
    
    // เติม UUID ถ้าขาด
    if (!uid) {
      uid = uuidMap[normalizeText(newN)];
    }
    
    // ตัดซ้ำ
    if (!seen[normOld]) {
      seen[normOld] = true;
      cleanList.push([oldN, newN, uid]);
    }
  });

  // 3. บันทึกกลับ
  if (cleanList.length > 0) {
    mapSheet.getRange(2, 1, mapSheet.getLastRow(), 3).clearContent();
    mapSheet.getRange(2, 1, cleanList.length, 3).setValues(cleanList);
    Browser.msgBox("✅ ซ่อมแซม Mapping เสร็จสิ้น (เหลือ " + cleanList.length + " รายการ)");
  }
}

// ==========================================
// 6. HELPER LOGIC (Full Clustering)
// ==========================================

function processClustering() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, 1, lastRow - 1, 15); // อ่านถึง Col O
  var values = range.getValues();
  
  var clusters = [];

  // Phase 1: หาตัวตั้งต้น (Verified Rows)
  values.forEach(function(r, idx) {
    if (r[CONFIG.COL_VERIFIED - 1] === true) {
      clusters.push({
        lat: parseFloat(r[CONFIG.COL_LAT - 1]),
        lng: parseFloat(r[CONFIG.COL_LNG - 1]),
        name: r[CONFIG.COL_SUGGESTED - 1] || r[CONFIG.COL_NAME - 1], // ใช้ชื่อ Suggested ก่อนถ้ามี
        rowIndexes: [idx],
        hasLock: true // ล็อกชื่อนี้ไว้เป็นแม่แบบ
      });
    }
  });

  // Phase 2: จับคู่แถวที่เหลือ (Unverified)
  values.forEach(function(r, idx) {
    if (r[CONFIG.COL_VERIFIED - 1] === true) return; // ข้ามพวก Verified ไปแล้ว

    var lat = parseFloat(r[CONFIG.COL_LAT - 1]);
    var lng = parseFloat(r[CONFIG.COL_LNG - 1]);
    
    if (isNaN(lat) || isNaN(lng)) return;

    var found = false;
    
    // วนลูปหา Cluster ที่ใกล้ที่สุด
    for (var c = 0; c < clusters.length; c++) {
      var dist = getHaversineDistanceKM(lat, lng, clusters[c].lat, clusters[c].lng);
      if (dist <= CONFIG.DISTANCE_THRESHOLD_KM) {
        clusters[c].rowIndexes.push(idx);
        found = true;
        break; // เจอแล้วหยุดหา (อยู่กลุ่มแรกที่เจอ)
      }
    }

    // ถ้าไม่เจอใครเลย ให้ตั้งตัวเป็น Cluster ใหม่
    if (!found) {
      clusters.push({
        lat: lat,
        lng: lng,
        rowIndexes: [idx],
        hasLock: false,
        name: null // ยังไม่มีชื่อชนะเลิศ
      });
    }
  });

  // Phase 3: ตัดสินชื่อผู้ชนะ (Best Name) และอัปเดตข้อมูล
  clusters.forEach(function(g) {
    var rawNames = g.rowIndexes.map(function(i) { return values[i][CONFIG.COL_NAME - 1]; });
    var winner = g.hasLock ? g.name : getBestName_Smart(rawNames); // ฟังก์ชันนี้อยู่ใน Utils_Common.gs
    var confidenceScore = g.rowIndexes.length;

    g.rowIndexes.forEach(function(idx) {
      // ถ้า Verified แล้วไม่ต้องแก้ชื่อ
      if (values[idx][CONFIG.COL_VERIFIED - 1] !== true) {
        values[idx][CONFIG.COL_SUGGESTED - 1] = winner;
        values[idx][CONFIG.COL_CONFIDENCE - 1] = confidenceScore;
        values[idx][CONFIG.COL_NORMALIZED - 1] = normalizeText(winner);
      }
    });
  });

  // บันทึกกลับลงชีต
  range.setValues(values);
  ss.toast("✅ จัดกลุ่มและแนะนำชื่อมาตรฐานเรียบร้อยแล้ว", "Clustering");
}
