/**
 * 📦 Service: SCG Operation (Final Integrated Version)
 * Version: 1.5 Final (Complete)
 * หน้าที่: 
// ==========================================
// 2. MAIN FUNCTIONS
// ==========================================

/**
 * 🚀 ฟังก์ชันหลัก: ดึงข้อมูลจาก SCG
 */
function fetchDataFromSCGJWD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const inputSheet = ss.getSheetByName(SCG_CONFIG.SHEET_INPUT);
    const dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
    if (!inputSheet || !dataSheet) throw new Error("ไม่พบชีต Input หรือ Data");

    // 1. ดึง Cookie และ Shipment
    const cookie = inputSheet.getRange(SCG_CONFIG.COOKIE_CELL).getValue();
    if (!cookie) throw new Error("ไม่พบ Cookie");

    const lastRow = inputSheet.getLastRow();
    if (lastRow < SCG_CONFIG.INPUT_START_ROW) throw new Error("ไม่พบ Shipment No.");

    const shipmentNumbers = inputSheet
      .getRange(SCG_CONFIG.INPUT_START_ROW, 1, lastRow - SCG_CONFIG.INPUT_START_ROW + 1, 1)
      .getValues().flat().filter(String);

    if (shipmentNumbers.length === 0) throw new Error("Shipment No. ว่างเปล่า");

    // แสดงผล Shipment ที่กำลังดึง
    const shipmentString = shipmentNumbers.join(',');
    inputSheet.getRange(SCG_CONFIG.SHIPMENT_STRING_CELL).setValue(shipmentString).setHorizontalAlignment("left");

    // 2. เรียก API
    const payload = {
      DeliveryDateFrom: '', DeliveryDateTo: '', TenderDateFrom: '', TenderDateTo: '',
      CarrierCode: '', CustomerCode: '', OriginCodes: '', ShipmentNos: shipmentString
    };
    
    const options = {
      method: 'post', payload: payload, muteHttpExceptions: true, headers: { cookie: cookie }
    };

    ss.toast("กำลังดึงข้อมูลจาก SCG...", "System", 60);
    const response = UrlFetchApp.fetch(SCG_CONFIG.API_URL, options); // *ตรวจสอบ URL ให้ตรงกับที่ท่านใช้จริง
    
    if (response.getResponseCode() !== 200) throw new Error("API Error: " + response.getContentText());
    
    // *หมายเหตุ: ถ้า API ของท่านคือ Link อื่น ให้แก้ตรง SCG_CONFIG.API_URL

    const json = JSON.parse(response.getContentText());
    const shipments = json.data || [];
    if (shipments.length === 0) throw new Error("ไม่พบข้อมูล Shipment");

    // 3. แปลงข้อมูล (Flatten)
    const allFlatData = [];
    let runningRow = 2;

    shipments.forEach(shipment => {
      // นับปลายทาง
      const destSet = new Set();
      (shipment.DeliveryNotes || []).forEach(n => { if (n.ShipToName) destSet.add(n.ShipToName); });
      const totalDestCount = destSet.size;
      const destListStr = Array.from(destSet).join(", ");

      (shipment.DeliveryNotes || []).forEach(note => {
        (note.Items || []).forEach(item => {
          const dailyJobId = note.PurchaseOrder + "-" + runningRow;
          
          // Row Structure (29 Columns)
          const row = [
            dailyJobId,                     // 0: ID
            note.PlanDelivery ? new Date(note.PlanDelivery) : null, // 1
            String(note.PurchaseOrder),     // 2
            String(shipment.ShipmentNo),    // 3
            shipment.DriverName,            // 4
            shipment.TruckLicense,          // 5
            String(shipment.CarrierCode),   // 6
            shipment.CarrierName,           // 7
            String(note.SoldToCode),        // 8
            note.SoldToName,                // 9
            note.ShipToName,                // 10
            note.ShipToAddress,             // 11
            note.ShipToLatitude + ", " + note.ShipToLongitude, // 12
            item.MaterialName,              // 13
            item.ItemQuantity,              // 14
            item.QuantityUnit,              // 15
            item.ItemWeight,                // 16
            String(note.DeliveryNo),        // 17
            totalDestCount,                 // 18
            destListStr,                    // 19
            "รอสแกน",                       // 20
            "ยังไม่ได้ส่ง",                   // 21
            "",                             // 22: Email Placeholder
            0, 0, 0,                        // 23-25: Aggregates
            "",                             // 26: LatLong Actual (รอเติม)
            "",                             // 27: Display Text
            shipment.ShipmentNo + "|" + note.ShipToName // 28: ShopKey
          ];
          allFlatData.push(row);
          runningRow++;
        });
      });
    });

    // 4. คำนวณยอดรวม (Aggregation)
    const shopAgg = {};
    allFlatData.forEach(r => {
      const key = r[28]; // ShopKey
      if (!shopAgg[key]) shopAgg[key] = { qty: 0, weight: 0, invoices: new Set(), epod: 0 };
      
      shopAgg[key].qty += Number(r[14]) || 0;
      shopAgg[key].weight += Number(r[16]) || 0;
      shopAgg[key].invoices.add(r[2]);
      if (checkIsEPOD(r[9], r[2])) shopAgg[key].epod++;
    });

    allFlatData.forEach(r => {
      const agg = shopAgg[r[28]];
      const scanInv = agg.invoices.size - agg.epod;
      r[23] = agg.qty;
      r[24] = Number(agg.weight.toFixed(2));
      r[25] = scanInv;
      r[27] = `${r[9]} / รวม ${scanInv} บิล`;
    });

    // 5. เขียนลงชีต
    const headers = [
      "ID_งานประจำวัน", "PlanDelivery", "InvoiceNo", "ShipmentNo", "DriverName",
      "TruckLicense", "CarrierCode", "CarrierName", "SoldToCode", "SoldToName",
      "ShipToName", "ShipToAddress", "LatLong_SCG", "MaterialName", "ItemQuantity", 
      "QuantityUnit", "ItemWeight", "DeliveryNo", "จำนวนปลายทาง_System", "รายชื่อปลายทาง_System", 
      "ScanStatus", "DeliveryStatus", "Email พนักงาน", 
      "จำนวนสินค้ารวมของร้านนี้", "น้ำหนักสินค้ารวมของร้านนี้", "จำนวน_Invoice_ที่ต้องสแกน",
      "LatLong_Actual", "ชื่อเจ้าของสินค้า_Invoice_ที่ต้องสแกน", "ShopKey"
    ];

    dataSheet.clear();
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    if (allFlatData.length > 0) {
      dataSheet.getRange(2, 1, allFlatData.length, headers.length).setValues(allFlatData);
      // Format Date
      dataSheet.getRange(2, 2, allFlatData.length, 1).setNumberFormat("dd/mm/yyyy");
      // Format Text for IDs
      dataSheet.getRange(2, 3, allFlatData.length, 1).setNumberFormat("@");
      dataSheet.getRange(2, 18, allFlatData.length, 1).setNumberFormat("@");
    }

    // 6. 🟢 เรียกฟังก์ชันจับคู่พิกัดทันที (ตัวที่ท่านถามถึง)
    applyMasterCoordinatesToDailyJob();
    
    ui.alert(`✅ ดึงข้อมูลสำเร็จ ${allFlatData.length} แถว และจับคู่พิกัดเรียบร้อย`);

  } catch (e) {
    ui.alert("❌ เกิดข้อผิดพลาด: " + e.message);
  }
}

/**
 * 🛰️ ฟังก์ชันจับคู่พิกัดและอีเมลพนักงาน (V1.2 Original Logic)
 * ถูกเรียกโดย: fetchDataFromSCGJWD (เมื่อดึงงาน) และ Agent (เมื่อซ่อมพิกัดเสร็จ)
 */
function applyMasterCoordinatesToDailyJob() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  const dbSheet = ss.getSheetByName(SCG_CONFIG.SHEET_MASTER_DB);
  const mapSheet = ss.getSheetByName(SCG_CONFIG.SHEET_MAPPING);
  const empSheet = ss.getSheetByName(SCG_CONFIG.SHEET_EMPLOYEE);

  if (!dataSheet || !dbSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return; // ไม่มีงาน ไม่ต้องทำ

  // 1. โหลด Master DB (ชื่อ -> พิกัด)
  const masterCoords = {};
  if (dbSheet.getLastRow() > 1) {
    // อ่าน Col 1(Name), 2(Lat), 3(Lng)
    dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 3).getValues().forEach(r => {
      if (r[0] && r[1] && r[2]) {
        masterCoords[normalizeText(r[0])] = r[1] + ", " + r[2];
      }
    });
  }

  // 2. โหลด Name Mapping (ชื่อเล่น -> ชื่อจริง)
  const aliasMap = {};
  if (mapSheet && mapSheet.getLastRow() > 1) {
    mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0] && r[1]) aliasMap[normalizeText(r[0])] = normalizeText(r[1]);
    });
  }

  // 3. โหลดข้อมูลพนักงาน (ชื่อคนขับ -> Email)
  const empMap = {};
  if (empSheet && empSheet.getLastRow() > 1) {
    empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues().forEach(r => {
      // Col B(1) = ชื่อ, Col G(6) = Email
      if (r[1] && r[6]) empMap[normalizeText(r[1])] = r[6];
    });
  }

  // 4. วนลูปงานในชีต Data เพื่ออัปเดต
  const range = dataSheet.getRange(2, 1, lastRow - 1, 29); // อ่านมาให้ครบ 29 Col
  const values = range.getValues();
  
  // เตรียม Array สำหรับเขียนกลับ (Performance Optimization)
  const latLongUpdates = [];
  const bgUpdates = [];
  const emailUpdates = [];

  values.forEach(r => {
    let newGeo = "";
    let bg = null;
    let email = r[22]; // ค่าเดิม

    // A. Map Coordinates
    // r[10] คือ ShipToName
    if (r[10]) { 
      let name = normalizeText(r[10]);
      
      // แปลงชื่อเล่นเป็นชื่อจริงก่อน
      if (aliasMap[name]) name = aliasMap[name];
      
      // หาพิกัดจาก Master
      if (masterCoords[name]) {
        newGeo = masterCoords[name];
        bg = "#b6d7a8"; // สีเขียว (เจอใน Master)
      } else {
        // ถ้าไม่เจอ ลองหาแบบสาขา (Branch Logic)
        // (ฟังก์ชันเสริม ถ้ามีใน Utils_Common)
        if (typeof findMasterByBranchLogic === 'function') {
             const byBranch = findMasterByBranchLogic(r[10], masterCoords);
             if (byBranch) { newGeo = byBranch; bg = "#b6d7a8"; }
        }
      }
    }
    latLongUpdates.push([newGeo]); // Col 27 (Index 26 ใน array นี้ แต่เขียนลง Col 27)
    bgUpdates.push([bg]);

    // B. Map Email
    // r[4] คือ DriverName
    if (r[4]) {
      const cleanDriver = normalizeText(r[4]);
      if (empMap[cleanDriver]) {
        email = empMap[cleanDriver];
      }
    }
    emailUpdates.push([email]);
  });

  // 5. บันทึกผลลัพธ์ทีเดียว (Batch Write)
  // Col 27 = LatLong_Actual
  dataSheet.getRange(2, 27, latLongUpdates.length, 1).setValues(latLongUpdates);
  dataSheet.getRange(2, 27, bgUpdates.length, 1).setBackgrounds(bgUpdates);
  
  // Col 23 = Email พนักงาน
  dataSheet.getRange(2, 23, emailUpdates.length, 1).setValues(emailUpdates);
}

/**
 * 🛠️ ฟังก์ชันล้างข้อมูลในชีต Data (Fix Bug: เช็คแถวก่อนลบ)
 */
function clearDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // ต้องมีข้อมูลมากกว่า 1 แถว (คือมีมากกว่าแค่ Header) ถึงจะลบได้
  if (lastRow > 1 && lastCol > 0) {
    const numRowsToDelete = lastRow - 1;
    sheet.getRange(2, 1, numRowsToDelete, lastCol).clearContent();
    sheet.getRange(2, 1, numRowsToDelete, lastCol).setBackground(null);
  }
}

/**
 * 🧹 ฟังก์ชันล้างข้อมูลทั้งหมด (เมนู)
 */
function clearAllSCGSheets() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('ยืนยันการล้างข้อมูล', 'คุณต้องการล้างข้อมูลในชีต Input และ Data ทั้งหมดหรือไม่?', ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ล้าง Input
    const inputSheet = ss.getSheetByName(SCG_CONFIG.SHEET_INPUT);
    if (inputSheet) {
      inputSheet.getRange(SCG_CONFIG.COOKIE_CELL).clearContent();
      inputSheet.getRange(SCG_CONFIG.SHIPMENT_STRING_CELL).clearContent();
      const lastRow = inputSheet.getLastRow();
      if (lastRow >= SCG_CONFIG.INPUT_START_ROW) {
        inputSheet.getRange(SCG_CONFIG.INPUT_START_ROW, 1, lastRow - SCG_CONFIG.INPUT_START_ROW + 1, 1).clearContent();
      }
    }

    // ล้าง Data
    clearDataSheet();

    ui.alert('✅ ล้างข้อมูลเรียบร้อย');
  }
}

// --- Helper Functions ---

function checkIsEPOD(ownerName, invoiceNo) {
  if (!ownerName || !invoiceNo) return false;
  const owner = String(ownerName).toUpperCase();
  const inv = String(invoiceNo).toUpperCase();
  const whitelist = ["SCG EXPRESS", "BETTERBE", "JWD TRANSPORT"];
  
  if (whitelist.some(w => owner.includes(w))) return true;
  if (["_DOC", "-DOC", "FFF", "EOP", "แก้เอกสาร"].some(k => inv.includes(k))) return false;
  if (inv.startsWith("N3")) return false;
  if (owner.includes("DENSO") || owner.includes("เด็นโซ่") || /^(78|79)/.test(inv)) return true;
  
  return false;
}

// Helper เผื่อไม่มีใน Utils_Common (ใส่กันไว้ก่อน)
function normalizeText(text) {
  if (!text) return "";
  return text.toString().toLowerCase().replace(/\s+/g, "").trim();
}
