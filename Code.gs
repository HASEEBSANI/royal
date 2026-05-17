
/***************************************************
 * GLOBAL SETTINGS
 ***************************************************/

const API_TOKEN = "MY_SECRET_123";

const SHEETS = {
  ACTIVATIONS_NAME: "Sheet1",
  REGISTRATIONS_NAME: "Form Responses",
  GIFT_NAME: "Sheet2",
  REQUIREMENTS_NAME: "Requirements"
};

const REG_SPREADSHEET_ID = "1nsSpRM4OtMur4YBTsF5TUd35D-CArRy9mKTDxsfNNUk";
const REQ_SPREADSHEET_ID = "1vP9dElDbDzMMm5b43jS31JDIG0a5FBn3MlbsE-Hxfys";
const GIF_SPREADSHEET_ID = "1YqMHnd6CSwTlXmdNTh327rOILxQs-Jwq0AOvVL0iFYU";
const ACT_SPREADSHEET_ID = "1de9MQaGXk0gM4gaO-hoCFHHgv1uuCnPutjzBNAvjSSw";
const PRICE_SPREADSHEET_ID = "12z48-52PTOa7ZPdB9Lm8J74LoPmH_P6373Xugj_N62Q";

const TIMEZONE = "GMT+5:30";

const PRICE_SHEETS = {
  MIPP: "Mobile",
  ECO: "ECO",
  TABLET: "Tablet",
  TV: "Tv"
};


/***************************************************
 * MAIN API ROUTER
 ***************************************************/

function doGet(e) {

  const action = e.parameter.action;

  if (action === "dashboard") {
    return outputJSON(getDashboardData());
  }

  if (action === "requirements") {
    return outputJSON(getRequirementData());
  }

  if (action === "recent") {
    return outputJSON(getRecentSubmissions());
  }

  if (action === "stats") {
    return outputJSON(getDashboardStats());
  }

  if (action === "retailers") {
    return outputJSON(getUniqueRetailers());
  }

  if (action === "stock") {
    return outputJSON(getStockData());
  }

  if (action === "giftEntries") {
    return outputJSON(getDashboardEntries());
  }

  if (action === "giftInit") {
    return outputJSON(getInitialData());
  }

  if (action === "priceList") {
    return outputJSON(getPriceList());
  }

  if (action === "mobileProducts") {
    return outputJSON(getProductsByMode("MIPP"));
  }

  if (action === "ecoProducts") {
    return outputJSON(getProductsByMode("ECO"));
  }

  if (action === "tabletProducts") {
    return outputJSON(getProductsByMode("TABLET"));
  }

  if (action === "tvProducts") {
    return outputJSON(getProductsByMode("TV"));
  }

  return outputJSON({
    status: "error",
    message: "Invalid action"
  });
}


function doPost(e) {

  try {

    const data = JSON.parse(e.postData.contents);

    if (data.token !== API_TOKEN) {
      return outputJSON({
        status: "error",
        message: "Unauthorized"
      });
    }

    switch (data.action) {

      case "processRequirement":
        return outputJSON(processRequirement(data.payload));

      case "processFormSubmission":
        return outputJSON(processFormSubmission(data.payload));

      case "processCSVUpload":
        return outputJSON(processCSVUpload(data.payload));

      case "newActivations":
        return outputJSON(processNewActivations(data.payload));

      case "deleteRequirement":
        return outputJSON(deleteRequirement(data.payload));

      default:
        return outputJSON({
          status: "error",
          message: "Unknown action"
        });
    }

  } catch(err) {

    return outputJSON({
      status: "error",
      message: err.toString()
    });

  }
}


function doOptions() {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}


function outputJSON(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/***************************************************
 * REQUIREMENTS
 ***************************************************/

function processRequirement(data) {

  const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);

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

  return {
    status: "success"
  };
}


function getRequirementData() {

  const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0];

  return data.slice(1).map((row, index) => {

    let obj = {
      _sheetRow: index + 2
    };

    headers.forEach((header, i) => {

      let val = row[i];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy HH:mm");
      }

      obj[header] = val;

    });

    return obj;

  }).reverse();
}


function deleteRequirement(rowNumber) {

  const ss = SpreadsheetApp.openById(REQ_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REQUIREMENTS_NAME);

  sheet.deleteRow(Number(rowNumber));

  return {
    status: "success"
  };
}


/***************************************************
 * DASHBOARD
 ***************************************************/

function getDashboardData() {

  const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0];

  return data.slice(1).map(row => {

    let obj = {};

    headers.forEach((header, i) => {

      let val = row[i];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy HH:mm");
      }

      obj[header] = val;

    });

    return obj;

  }).reverse();
}


/***************************************************
 * RECENT REGISTRATIONS
 ***************************************************/

function getRecentSubmissions() {

  const ss = SpreadsheetApp.openById(REG_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS_NAME);

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0];

  return data.slice(1).map(row => {

    let obj = {};

    headers.forEach((header, i) => {

      let val = row[i];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, TIMEZONE, "dd-MM-yyyy");
      }

      obj[header] = val;

    });

    return obj;

  }).reverse();
}


/***************************************************
 * ACTIVATION CSV UPLOAD
 ***************************************************/

function processCSVUpload(csvData) {

  const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

  let sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.ACTIVATIONS_NAME);
  }

  const csvRows = Utilities.parseCsv(csvData);

  if (csvRows.length < 2) {
    return {
      status: "error",
      message: "CSV Empty"
    };
  }

  const dataOnly = csvRows.slice(1);

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      dataOnly.length,
      dataOnly[0].length
    )
    .setValues(dataOnly);

  return {
    status: "success",
    count: dataOnly.length
  };
}


function processNewActivations(dataRows) {

  const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

  const existingData = sheet.getDataRange().getValues();

  const existingIMEIs = existingData.map(r => String(r[0]).trim());

  const rowsToAppend = [];

  dataRows.forEach(rowStr => {

    const cols = rowStr.split('\t');

    const imei = cols[0].trim();

    if (existingIMEIs.indexOf(imei) === -1) {
      rowsToAppend.push(cols);
    }

  });

  if (rowsToAppend.length > 0) {

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      )
      .setValues(rowsToAppend);

    return {
      status: "success",
      count: rowsToAppend.length
    };
  }

  return {
    status: "error",
    message: "Duplicates only"
  };
}


/***************************************************
 * PRICE LIST
 ***************************************************/

function getProductsByMode(mode) {

  const sheetName = PRICE_SHEETS[mode];

  const ss = SpreadsheetApp.openById(PRICE_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0];

  return data.slice(1).map(row => {

    let obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    return obj;

  });
}


function getPriceList() {
  return getProductsByMode("MIPP");
}


/***************************************************
 * GIFT FORM
 ***************************************************/

function getInitialData() {

  const ss = SpreadsheetApp.openById(GIF_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Sheet1");

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return {
      dmsData: [],
      handsets: [],
      prizes: []
    };
  }

  const rows = values.slice(1);

  const dmsMap = new Map();
  const handsetSet = new Set();
  const prizeSet = new Set();

  rows.forEach(row => {

    const dmsCode = String(row[0] || "").trim();
    const retailerName = String(row[1] || "").trim();
    const handset = String(row[2] || "").trim();
    const prize = String(row[3] || "").trim();

    if (dmsCode && retailerName) {
      dmsMap.set(dmsCode, {
        dmsCode,
        retailerName
      });
    }

    if (handset) handsetSet.add(handset);
    if (prize) prizeSet.add(prize);

  });

  return {
    dmsData: Array.from(dmsMap.values()),
    handsets: Array.from(handsetSet),
    prizes: Array.from(prizeSet)
  };
}


function processFormSubmission(data) {

  const DRIVE_FOLDER_ID = "1ji5VfBPYt2ZmLjOW_BIkJ-YPujFa3XtS";

  const ss = SpreadsheetApp.openById(GIF_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Sheet2");

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  const rowData = [
    new Date(),
    data.dmsCode,
    data.retailerName,
    data.customerName,
    data.mobileNo,
    data.purchasedHandset,
    data.giftPrizeWonBy
  ];

  const fileFields = [
    "purchasedBill",
    "handoverPhoto",
    "giftPrizeScreenshot",
    "idProof"
  ];

  fileFields.forEach(field => {

    const fileData = data[field];

    if (fileData && fileData.base64Data) {

      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileData.base64Data),
        fileData.mimeType,
        fileData.name
      );

      const file = folder.createFile(blob);

      rowData.push(file.getUrl());

    } else {

      rowData.push("File Missing");

    }

  });

  sheet.appendRow(rowData);

  return {
    status: "success"
  };
}
```

# GitHub Frontend Changes

```javascript
const API_URL = "YOUR_WEBAPP_URL";
const API_TOKEN = "MY_SECRET_123";
```

# Example GET

```javascript
fetch(`${API_URL}?action=dashboard`)
.then(r => r.json())
.then(data => console.log(data));
```

# Example POST

```javascript
fetch(API_URL, {
  method: "POST",
  headers: {
    "Content-Type": "application/json"
  },
  body: JSON.stringify({
    token: API_TOKEN,
    action: "processRequirement",
    payload: formData
  })
})
.then(r => r.json())
.then(data => console.log(data));
```

# Deploy Settings

1. Deploy → New Deployment
2. Type → Web App
3. Execute as → Me
4. Access → Anyone
5. Deploy
6. Copy Web App URL
7. Paste inside GitHub HTML/JS
