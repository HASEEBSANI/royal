/***************************************************
 * 1. GLOBAL SETTINGS & ROUTING
 ***************************************************/
const SHEETS = {
  ACTIVATIONS_NAME: "Sheet1",        
  REGISTRATIONS_NAME: "Form Responses", 
  GIFT_NAME: "Sheet2",
  REQUIREMENTS_NAME: "Requirements" 
};

// EXTERNAL SPREADSHEET IDs
const REG_SPREADSHEET_ID = "1nsSpRM4OtMur4YBTsF5TUd35D-CArRy9mKTDxsfNNUk"; 
const REQ_SPREADSHEET_ID = "1vP9dElDbDzMMm5b43jS31JDIG0a5FBn3MlbsE-Hxfys";
const GIF_SPREADSHEET_ID = "1YqMHnd6CSwTlXmdNTh327rOILxQs-Jwq0AOvVL0iFYU";
const ACT_SPREADSHEET_ID = "1de9MQaGXk0gM4gaO-hoCFHHgv1uuCnPutjzBNAvjSSw";
const PRICE_SPREADSHEET_ID = "12z48-52PTOa7ZPdB9Lm8J74LoPmH_P6373Xugj_N62Q";

const TIMEZONE = "GMT+5:30";

/***************************************************
 * PRICE LIST SETTINGS
 ***************************************************/
const PRICE_SHEETS = {
  MIPP: "Mobile",
  ECO: "ECO",
  TABLET: "Tablet",
  TV: "Tv"
};

function doGet(e) {
  let page = (e && e.parameter && e.parameter.p) ? e.parameter.p : "index";
  let fileName = page.toLowerCase(); 
  try {
    return HtmlService.createTemplateFromFile(fileName)
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle("Activation System | " + page.toUpperCase());
  } catch (err) {
    return HtmlService.createHtmlOutput("<b>Error:</b> The file '" + fileName + ".html' was not found.");
  }
}

function getScriptURL() {
  return ScriptApp.getService().getUrl();
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
/***************************************************
 * 2. REQUIREMENT HISTORY LOGIC
 ***************************************************/
function getRequirementData() {
  try {
    const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);
    if (!sheet) return { rows: [], totalCount: 0, headers: [] };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) return { rows: [], totalCount: 0, headers: [] }; 

    const headers = data[0]; // Get spreadsheet column headers
    
    const rows = data.slice(1).map((row, rowIndex) => {
      let obj = {};
      obj["_sheetRow"] = rowIndex + 2; // Store physical row for deletion

      headers.forEach((header, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy HH:mm");
        }
        obj[header] = (val === "" || val === null) ? "-" : val;
      });
      return obj;
    }).reverse(); 
    
    return { 
      rows: rows, 
      headers: headers,
      totalCount: rows.length,
      loginTime: Utilities.formatDate(new Date(), TIMEZONE, "hh:mm a")
    };
  } catch (e) {
    return { rows: [], totalCount: 0, error: e.toString() };
  }
}

/***************************************************
 * 4. REGISTRATION LOGIC (ID: xsfNNUk)
 ***************************************************/

function getRecentSubmissions() {
  try {
    const ss = SpreadsheetApp.openById(REG_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS_NAME);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        if (val instanceof Date) val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy");
        
        // Handle Drive links
        if (header.toString().toUpperCase().includes("URL") || header.toString().toUpperCase().includes("FILE")) {
          val = convertToDirectDownload(val);
        }
        obj[header] = val;
      });
      return obj;
    }).reverse();
  } catch (e) { return []; }
}

/***************************************************
 * 5. ACTIVATION UPLOAD LOGIC
 ***************************************************/
/**
 * Fetches raw data from the Activation Spreadsheet for the Activity Table
 */
function getStockData() {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);
    if (!sheet) return [];
    
    const values = sheet.getDataRange().getValues();
    
    // Format dates to string before sending to frontend to avoid "disorderly" JS Date objects
    const formattedValues = values.map(row => row.map(cell => {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, TIMEZONE, "dd-MM-yyyy");
      }
      return cell;
    }));
    
    return formattedValues;
  } catch (e) {
    return [];
  }
}


function processCSVUpload(csvData) {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME) || ss.insertSheet(SHEETS.ACTIVATIONS_NAME);
    const csvRows = Utilities.parseCsv(csvData);
    
    if (csvRows.length < 2) return "Error: CSV is empty";
    
    const dataOnly = csvRows.slice(1);
    sheet.getRange(sheet.getLastRow() + 1, 1, dataOnly.length, dataOnly[0].length).setValues(dataOnly);
    return "Successfully imported " + dataOnly.length + " rows!";
  } catch (e) { return "Error: " + e.toString(); }
}

function processNewActivations(dataRows) {
  const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);
  const existingData = sheet.getDataRange().getValues();
  const existingIMEIs = existingData.map(r => String(r[0]).trim()); // Column A
  
  const rowsToAppend = [];
  
  dataRows.forEach(rowStr => {
    const cols = rowStr.split('\t');
    const imei = cols[0].trim();
    
    // Final server-side check
    if (existingIMEIs.indexOf(imei) === -1) {
      rowsToAppend.push(cols);
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    return { status: "success", count: rowsToAppend.length };
  } else {
    throw "All entries were duplicates!";
  }
}

/***************************************************
 * 6. UTILITIES
 ***************************************************/

function convertToDirectDownload(url) {
  if (!url || typeof url !== 'string' || !url.includes("drive.google.com")) return url;
  const match = url.match(/[-\w]{25,}/);
  return match ? "https://drive.google.com/uc?export=download&id=" + match[0] : url;
}

/***************************************************
 * 7. UPLOAD HISTORY
 ***************************************************/

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);
    if (!sheet) return { rows: [], totalCount: 0 };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { rows: [], totalCount: 0 }; 

    const headers = data[0];
    
    const rows = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        let h = header.toString().trim();

        // Specific Mapping: Map Sheet Headers to HTML Table Variables
        let key = "";
        if (h === "IMEI") key = "imei";
        else if (h === "Product Name") key = "product";
        else if (h === "Acivated Time") key = "activated"; // Matching your spelling "Acivated"
        else if (h === "DMS Code") key = "dms";
        else if (h === "Retailer Name") key = "retailer";
        else key = h.toLowerCase().replace(/\s+/g, ''); // Fallback for other columns

        // Handle Date/Time Formatting
        if (val instanceof Date) {
          // Format as DD-MM-YYYY HH:mm
          val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy HH:mm");
        }
        
        obj[key] = (val === "" || val === null) ? "-" : val;
      });
      return obj;
    }).reverse(); 
    
    return { 
      rows: rows, 
      totalCount: rows.length,
      loginTime: Utilities.formatDate(new Date(), TIMEZONE, "hh:mm a")
    };
  } catch (e) {
    return { rows: [], error: e.toString() };
  }
}

/***************************************************
 * 8. GIFT SUBMISSION
 ***************************************************/
function getDashboardEntries() {
  try {
    const ss = SpreadsheetApp.openById(GIF_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.GIFT_NAME);
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();

    // If only header exists or sheet is empty
    if (values.length < 2) {
      return [];
    }

    const data = values.slice(1).map(row => ({
      dms: row[1] || "N/A",
      retailer: row[2] || "N/A",
      customerName: row[3] || "N/A",
      mobileNo: row[4] || "",
      handset: row[5] || "N/A",
      prizeWon: row[6] || "N/A",
      billUrl: row[7] || "",
      handoverUrl: row[8] || "",
      screenshotUrl: row[9] || "",
      idProofUrl: row[10] || ""
    })).filter(item => item.customerName !== "N/A");

    return data;

  } catch (e) {
    return [];
  }
}

/***************************************************
 * 9. GIFT
 ***************************************************/
/**
 * Loads initial data (DMS Codes, Handsets, Prizes) from 'Sheet1'.
 * Renamed to getInitialData to match your HTML's call.
 */
function getInitialData() {
  const ss = SpreadsheetApp.openById(GIF_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Sheet1'); 
  
  if (!sheet) {
    throw new Error("Sheet named 'Sheet1' not found.");
  }
  
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { dmsData: [], handsets: [], prizes: [] };
  }
  
  const dataRows = values.slice(1);
  const dmsMap = new Map();
  const handsetSet = new Set();
  const prizeSet = new Set();

  dataRows.forEach(row => {
    const dmsCode = String(row[0] || '').trim();
    const retailerName = String(row[1] || '').trim();
    const handset = String(row[2] || '').trim();
    const prize = String(row[3] || '').trim();
    
    if (dmsCode && retailerName) {
      dmsMap.set(dmsCode, { dmsCode: dmsCode, retailerName: retailerName });
    }
    if (handset) handsetSet.add(handset);
    if (prize) prizeSet.add(prize);
  });
  
  return {
    dmsData: Array.from(dmsMap.values()),
    handsets: Array.from(handsetSet).sort(),
    prizes: Array.from(prizeSet).sort()
  };
}

/**
 * Processes form submission and saves to 'Sheet2'
 */
function processFormSubmission(data) {
  const DRIVE_FOLDER_ID = '1ji5VfBPYt2ZmLjOW_BIkJ-YPujFa3XtS'; 
  const ss = SpreadsheetApp.openById(GIF_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Sheet2'); 
  
  if (!sheet) throw new Error("Sheet 'Sheet2' not found.");

  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    throw new Error("Drive Folder access denied. Check Folder ID and permissions.");
  }
  
  // 1. Prepare Base Data (Columns A to G)
  const rowData = [
    new Date(), 
    data.dmsCode, 
    data.retailerName, 
    data.customerName, 
    data.mobileNo, 
    data.purchasedHandset, 
    data.giftPrizeWonBy 
  ];
  
  // 2. Process Files (Columns H to K)
  const fileFields = ['purchasedBill', 'handoverPhoto', 'giftPrizeScreenshot', 'idProof'];
  
  fileFields.forEach(field => {
    const fileData = data[field];
    if (fileData && fileData.base64Data) {
      const fileExtension = fileData.name.split('.').pop();
      const safeCustomerName = data.customerName.replace(/[^a-z0-9]/gi, '_');
      const fileName = `${data.dmsCode}_${safeCustomerName}_${field}.${fileExtension}`;
      
      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileData.base64Data), 
        fileData.mimeType, 
        fileName
      );
      
      const file = folder.createFile(blob);
      rowData.push(file.getUrl());
    } else {
      rowData.push('File Missing');
    }
  });

  sheet.appendRow(rowData);
  return "Submission successful"; 
}


/***************************************************
 * 10. RETAILER (MASTER LOADER & OVERVIEW)
 ***************************************************/

function getUniqueRetailers() {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Sheet1"); // Ensure this matches your sheet name
    if (!sheet) return ["Error: Sheet1 not found"];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return ["No Data Available"];

    const headers = data[0];
    const dmsIdx = headers.findIndex(h => /dms/i.test(h.toString()));
    const nameIdx = headers.findIndex(h => /retailer\s*name/i.test(h.toString()));

    const uniqueSet = new Set();
    for (let i = 1; i < data.length; i++) {
      const dms = String(data[i][dmsIdx] || "").trim();
      const name = String(data[i][nameIdx] || "").trim();
      if (dms && name) {
        uniqueSet.add(dms + " - " + name);
      }
    }

    const result = Array.from(uniqueSet);

    // SORT BY NAME ONLY
    result.sort((a, b) => {
      const nameA = a.includes(" - ") ? a.split(" - ")[1].trim() : a;
      const nameB = b.includes(" - ") ? b.split(" - ")[1].trim() : b;
      return nameA.localeCompare(nameB, undefined, {
        numeric: true,
        sensitivity: 'base'
      });
    });

    return result;

  } catch (err) {
    return ["ERROR: " + err.toString()];
  }
}

function getSpecificRetailerData(retailerName) {
  const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID); // ✅ your sheet ID
  const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

  if (!sheet) return {
    yesterday: 0,
    week: 0,
    month: 0,
    lastMonth: 0,
    months: [],
    storeList: []
  };

  // ✅ Get all data
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return {
      yesterday: 0,
      week: 0,
      month: 0,
      lastMonth: 0,
      months: [],
      storeList: []
    };
  }

  // ✅ Headers from first row
  const headers = data.shift(); // first row as header

  // 🔎 Auto detect date & retailer columns
  const dateIndex = headers.findIndex(h =>
    /date|time|activ|aciv/i.test(String(h))
  );
  const retailerIndex = headers.findIndex(h =>
    /retailer/i.test(String(h))
  );

  if (dateIndex === -1 || retailerIndex === -1) {
    throw new Error("Date or Retailer column not found. Check header spelling.");
  }

  // ✅ Date ranges
  const today = new Date();
  today.setHours(0,0,0,0);

  const yesterdayDate = new Date(today);
  yesterdayDate.setDate(today.getDate() - 1);

  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - 6);

  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const startOfLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const endOfLastMonth = new Date(today.getFullYear(), today.getMonth(), 0);

  let yesterday = 0;
  let week = 0;
  let month = 0;
  let lastMonth = 0;

  const storeMap = {};

  // ✅ Loop through rows
  data.forEach(row => {
    const rawDate = row[dateIndex];
    const rowRetailer = row[retailerIndex];
    if (!rawDate || !rowRetailer) return;

    // Clean retailer match
    let selectedName = retailerName;
    if (retailerName && retailerName.includes(" - ")) {
      selectedName = retailerName.split(" - ")[1];
    }
    const cleanSelected = selectedName ? selectedName.toString().trim().toLowerCase() : "";
    const cleanRowRetailer = rowRetailer.toString().trim().toLowerCase();
    if (cleanSelected && cleanSelected !== cleanRowRetailer) return;

    // Convert date properly
    let rowDate;
    if (rawDate instanceof Date) {
      rowDate = rawDate;
    } else {
      const fixed = rawDate.toString().replace(" ", "T");
      rowDate = new Date(fixed);
    }
    if (!rowDate || isNaN(rowDate.getTime())) return;

    rowDate.setHours(0,0,0,0);

    // ✅ Count metrics
    if (isSameDate(rowDate, yesterdayDate)) yesterday++;
    if (rowDate >= startOfWeek) week++;
    if (rowDate >= startOfMonth) month++;
    if (rowDate >= startOfLastMonth && rowDate <= endOfLastMonth) lastMonth++;

    const monthKey = rowDate.getFullYear() + "-" + (rowDate.getMonth() + 1);

    if (!storeMap[rowRetailer]) {
      storeMap[rowRetailer] = { name: rowRetailer, months: {} };
    }
    if (!storeMap[rowRetailer].months[monthKey]) {
      storeMap[rowRetailer].months[monthKey] = 0;
    }
    storeMap[rowRetailer].months[monthKey]++;
  });

  // ✅ Prepare last 6 months for table header
  const months = [];
for (let i = 0; i < 6; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    months.push({
        key: d.getFullYear() + "-" + (d.getMonth() + 1),
        label: Utilities.formatDate(d, TIMEZONE, "MMM yy")
    });
}

const storeList = Object.values(storeMap);

// ✅ Sort retailer names alphabetically A-Z
storeList.sort((a, b) => {
  return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
});

return {
    yesterday,
    week,
    month,
    lastMonth,
    months,
    storeList
};
}

// ✅ Helper function
function isSameDate(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}




/***************************************************
 * 11. PRICE LIST
 ***************************************************/
function getProductsByMode(mode) {
  try {
    const sheetName = PRICE_SHEETS[mode];
    if (!sheetName) return [];

    const ss = SpreadsheetApp.openById(PRICE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];

    const headersRaw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Normalize headers: Billing → NLC
    const headers = headersRaw.map(h => {
      if (!h) return "";
      let hh = String(h).trim();
      if (hh.toLowerCase() === "billing") hh = "NLC";
      return hh;
    });

    const products = rows.map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = r[i];

        // ✅ Fix floating point issue (7599.049999 → 7599.04)
        if (typeof val === "number") {
  val = Number(val.toFixed(2));
}

        obj[h] = val;
      });
      return obj;
    });

    return products.filter(p => {
      const pn = p["Product Name"] || p["product name"] || p["Product"];
      return pn && String(pn).trim() !== "";
    });

  } catch (e) {
    return [];
  }
}

function getPriceList() {
  const sheet = SpreadsheetApp.openById("12z48-52PTOa7ZPdB9Lm8J74LoPmH_P6373Xugj_N62Q")
    .getSheetByName("mobile");

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveProductToSheet(product, index) {

  const sheet = SpreadsheetApp.openById("12z48-52PTOa7ZPdB9Lm8J74LoPmH_P6373Xugj_N62Q")
    .getSheetByName("mobile");

  const headers = sheet.getRange(1,1,1,5).getValues()[0];

  const rowData = [
    product.name,
    product.billing,
    product.dp,
    product.srp,
    product.mrp
  ];

  if (index === "" || index === null) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(Number(index) + 2, 1, 1, 5).setValues([rowData]);
  }

  return true;
}


/***************************************************
 * REQUIREMENT FORM SUBMISSION
 ***************************************************/
function processRequirement(data) {
  try {
    const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);
    if (!sheet) throw new Error("Requirements sheet not found");

    const rowData = [
      new Date(),
      data.retailerName || "",
      data.redmi14c_4_64 || 0,
      data.redmi14c_4_128 || 0,
      data.redmi14c_6_128 || 0,
      data.note14_6_128 || 0,
      data.note14_8_128 || 0,
      data.note14_8_256 || 0,
      data.note14pro_8_128 || 0,
      data.note14pro_8_256 || 0,
      data.note14proPlus_8_256 || 0,
      data.note14pro_12_512 || 0
    ];

    sheet.appendRow(rowData);

    return "SUCCESS";
  } catch (e) {
    throw new Error(e.toString());
  }
}

/***************************************************
 * DELETE REQUIREMENT ROW
 ***************************************************/
function deleteRequirement(rowNumber) {
  try {
    const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);

    if (!sheet) throw new Error("Requirements sheet not found");

    sheet.deleteRow(Number(rowNumber)); // 🔥 NO +1

    return "SUCCESS";
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

/***************************************************
 * 11. PRICE LIST (CLEAN VERSION)
 ***************************************************/
function getPriceList() {
  const sheet = SpreadsheetApp.openById(PRICE_SPREADSHEET_ID)
    .getSheetByName("mobile");

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data.shift();

  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveProductToSheet(product, index) {

  const sheet = SpreadsheetApp.openById(PRICE_SPREADSHEET_ID)
    .getSheetByName("mobile");

  const rowData = [
    product.name,
    product.billing,
    product.dp,
    product.srp,
    product.mrp
  ];

  if (index === "" || index === null) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(Number(index) + 2, 1, 1, 5).setValues([rowData]);
  }

  return true;
}

/***************************************************
 * 12. PERFORMANCE OVERVIEW
 ***************************************************/
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);
    if (!sheet) return defaultStats();

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return defaultStats();

    const headers = data.shift();

    const dateIndex = headers.findIndex(h =>
      /date|time|activ|aciv/i.test(String(h))
    );

    const retailerIndex = headers.findIndex(h =>
      /retailer/i.test(String(h))
    );

    if (dateIndex === -1 || retailerIndex === -1) {
      return defaultStats();
    }

    const today = new Date();
    const todayDate = today.getDate();

    const startOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);

    const lastMonthDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const lastMonth = lastMonthDate.getMonth();
    const lastMonthYear = lastMonthDate.getFullYear();

    let thisMonthCount = 0;
    let lastMonthCount = 0;

    const activeRetailerSet = new Set();   // This month active
    const totalRetailerSet = new Set();    // All retailers
    const retailerMap = {};

    data.forEach(row => {
      const rawDate = row[dateIndex];
      const retailer = row[retailerIndex];
      if (!rawDate || !retailer) return;

      const rowDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(rowDate)) return;

      const name = retailer.toString().trim();

      // Collect total retailer base
      totalRetailerSet.add(name);

      // THIS MONTH (1 → today)
      if (rowDate >= startOfThisMonth) {
        thisMonthCount++;
        activeRetailerSet.add(name);
        retailerMap[name] = (retailerMap[name] || 0) + 1;
      }

      // LAST MONTH TILL SAME DATE
      if (
        rowDate.getMonth() === lastMonth &&
        rowDate.getFullYear() === lastMonthYear &&
        rowDate.getDate() <= todayDate
      ) {
        lastMonthCount++;
      }
    });

    // Growth %
    let growthPercent = 0;
    if (lastMonthCount > 0) {
      growthPercent =
        ((thisMonthCount - lastMonthCount) / lastMonthCount) * 100;
    }

    const sortedRetailers = Object.entries(retailerMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 7);

    return {
      totalCount: thisMonthCount,
      lastMonthCount,
      growthPercent: growthPercent.toFixed(1),
      outletActive: activeRetailerSet.size,
      totalRetailers: totalRetailerSet.size,
      trend: sortedRetailers
    };

  } catch (e) {
    return defaultStats();
  }
}

function defaultStats() {
  return {
    totalCount: 0,
    lastMonthCount: 0,
    growthPercent: 0,
    outletActive: 0,
    totalRetailers: 0,
    trend: []
  };
}


/***************************************************
 * 14. RETAILER UPLOADS
 ***************************************************/
function uploadTargetCSV(csvData) {
  try {
    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Target");
    if (!sheet) sheet = ss.insertSheet("Target");
    else sheet.clear();

    const rows = Utilities.parseCsv(csvData);
    if (rows.length === 0) return { status: "error", message: "CSV empty" };

    // Optional: validate header
    if (rows[0][0] !== "DMS CODE" || rows[0][1] !== "RETAILER NAME" || rows[0][2] !== "TARGET") {
      return { status: "error", message: "Invalid CSV header" };
    }

    const header = rows.shift();
    rows.sort((a, b) => a[1].localeCompare(b[1])); // sort by retailer
    rows.unshift(header);

    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    return { status: "success", message: `CSV uploaded successfully! ${rows.length - 1} records.` };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
function getTargetTemplate() {
  const headers = ["DMS CODE", "RETAILER NAME", "TARGET"];
  return headers.join(",") + "\n"; // Returns CSV headers
}

/***************************************************
 * 15. FOS MAPPING
 ***************************************************/

 const SPREADSHEET_ID = '1de9MQaGXk0gM4gaO-hoCFHHgv1uuCnPutjzBNAvjSSw';
const SHEET_NAME = 'FOS';

function uploadTargetCSV(csvData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["DMS Code", "FOS Name"]);
    }

    const rows = Utilities.parseCsv(csvData);
    if (rows.length < 2) throw new Error("CSV is empty or missing data.");

    // Extract the FOS Name from the first data row of the upload
    // Assumes one FOS upload at a time. 
    const targetFOS = rows[1][1]; 

    const fullData = sheet.getDataRange().getValues();
    const headers = fullData[0];
    
    // Filter out existing rows that match the uploaded FOS Name
    // This effectively "deletes" only that FOS's data to make room for the new list
    const remainingData = fullData.filter((row, index) => {
      if (index === 0) return false; // remove header to re-add later
      return row[1] !== targetFOS; 
    });

    // Combine remaining data with the new CSV data (excluding CSV header)
    const newDataToInput = rows.slice(1);
    const finalData = [headers, ...remainingData, ...newDataToInput];

    // Clear and Rewrite
    sheet.clearContents();
    sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);

    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getTargetTemplate() {
  // Returns the header for the CSV download
  return "DMS Code,FOS Name";
}

function downloadTemplate(){
  google.script.run.withSuccessHandler(csv=>{
    const blob=new Blob([csv], {type:'text/csv'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; 
    a.download="fos_mapping_template.csv"; // Updated name
    a.click();
    URL.revokeObjectURL(url);
  }).getTargetTemplate();
}
