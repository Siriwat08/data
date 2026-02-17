/**
 * 🔍 Service: Search API (V3.0 - Pagination & Agent Support)
 * หน้าที่: ค้นหาข้อมูลแบบแบ่งหน้า (Google Style) + รองรับข้อมูลจาก Agent
 */

function searchMasterData(keyword, page) {
  var pageNum = page || 1;
  var pageSize = 20;

  if (!keyword || keyword.trim() === "") return { items: [], total: 0, totalPages: 0 };
  
  var rawKey = keyword.trim().toLowerCase();
  var searchKey = normalizeText(keyword); 

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. โหลด NameMapping
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET); 
  var aliasMap = {}; 
  if (mapSheet && mapSheet.getLastRow() > 1) {
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
    mapData.forEach(function(row) {
      var alias = row[0], master = row[1];
      if (alias && master) {
        var cleanMaster = normalizeText(master);
        if (!aliasMap[cleanMaster]) aliasMap[cleanMaster] = "";
        aliasMap[cleanMaster] += " " + normalizeText(alias) + " " + alias.toString().toLowerCase();
      }
    });
  }

  // 2. ค้นหาใน Database
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { items: [], total: 0, totalPages: 0 };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [], total: 0, totalPages: 0 };

  // อ่านข้อมูล Col A ถึง Col Q (หรือตาม Config)
  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); 
  var matches = []; 

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = row[CONFIG.COL_NAME - 1];
    if (!name) continue;

    var address = row[CONFIG.COL_ADDR_GOOG - 1] || row[CONFIG.COL_SYS_ADDR - 1];
    var lat = row[CONFIG.COL_LAT - 1];
    var lng = row[CONFIG.COL_LNG - 1];
    var uuid = row[CONFIG.COL_UUID - 1];
    
    // ✅ ดึงสิ่งที่ Agent เขียนไว้ (Col F) มาใช้ค้นหาด้วย
    var aiKeywords = row[CONFIG.COL_NORMALIZED - 1] ? row[CONFIG.COL_NORMALIZED - 1].toString().toLowerCase() : "";

    var normName = normalizeText(name);
    var normAddr = address ? normalizeText(address) : "";
    var rawName = name.toString().toLowerCase();
    var aliases = aliasMap[normName] || "";

    if (
      normName.includes(searchKey) || 
      rawName.includes(rawKey) ||
      aliases.includes(searchKey) || 
      aliases.includes(rawKey) ||
      aiKeywords.includes(searchKey) || // 👈 จุดสำคัญ: ค้นหาในสมอง Agent
      aiKeywords.includes(rawKey) ||    
      normAddr.includes(searchKey)
    ) {
      matches.push({
        name: name,
        address: address,
        lat: lat,
        lng: lng,
        mapLink: (lat && lng) ? "https://www.google.com/maps/dir/?api=1&destination=" + lat + "," + lng : "",
        uuid: uuid
      });
    }
  }

  // 3. ตัดแบ่งหน้า (Pagination Logic)
  var totalItems = matches.length;
  var totalPages = Math.ceil(totalItems / pageSize);
  
  if (pageNum > totalPages) pageNum = 1;
  
  var startIndex = (pageNum - 1) * pageSize;
  var endIndex = startIndex + pageSize;
  var pagedItems = matches.slice(startIndex, endIndex);

  return {
    items: pagedItems,
    total: totalItems,
    totalPages: totalPages,
    currentPage: pageNum
  };
}
