/***************************************************
 * COMPLETE CODE.GS
 ***************************************************/

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

/***************************************************
 * 2. ROUTING
 ***************************************************/
function doGet(e) {

  let page = (e && e.parameter && e.parameter.p)
    ? e.parameter.p
    : "index";

  let fileName = page.toLowerCase();

  try {

    return HtmlService
      .createTemplateFromFile(fileName)
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle("Activation System | " + page.toUpperCase());

  } catch (err) {

    return HtmlService.createHtmlOutput(
      "<h2>ERROR</h2><p>File not found : " + fileName + ".html</p>"
    );

  }

}

function getScriptURL() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}

/***************************************************
 * 3. DASHBOARD DATA
 ***************************************************/
function getDashboardData() {

  try {

    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

    if (!sheet) {
      return {
        rows: [],
        totalCount: 0,
        error: "Sheet1 not found"
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return {
        rows: [],
        totalCount: 0
      };
    }

    const headers = data[0];

    const rows = data.slice(1).map(row => {

      let obj = {};

      headers.forEach((header, i) => {

        let val = row[i];
        let h = String(header).trim();

        let key = "";

        // HEADER MAPPING
        if (/imei/i.test(h)) {
          key = "imei";
        }

        else if (/product/i.test(h)) {
          key = "product";
        }

        else if (/activ|aciv/i.test(h)) {
          key = "activated";
        }

        else if (/dms/i.test(h)) {
          key = "dms";
        }

        else if (/retailer/i.test(h)) {
          key = "retailer";
        }

        else {
          key = h.toLowerCase().replace(/\s+/g, '');
        }

        // DATE FORMAT
        if (val instanceof Date) {

          val = Utilities.formatDate(
            val,
            TIMEZONE,
            "dd-MM-yyyy HH:mm"
          );

        }

        obj[key] =
          (val === "" || val === null)
            ? "-"
            : val;

      });

      return obj;

    }).reverse();

    return {
      rows: rows,
      totalCount: rows.length,
      loginTime: Utilities.formatDate(
        new Date(),
        TIMEZONE,
        "hh:mm a"
      )
    };

  } catch (e) {

    return {
      rows: [],
      totalCount: 0,
      error: e.toString()
    };

  }

}

/***************************************************
 * 4. STOCK DATA
 ***************************************************/
function getStockData() {

  try {

    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    const sheet = ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();

    return values.map(row =>
      row.map(cell => {

        if (cell instanceof Date) {

          return Utilities.formatDate(
            cell,
            TIMEZONE,
            "dd-MM-yyyy HH:mm"
          );

        }

        return cell;

      })
    );

  } catch (e) {

    return [];

  }

}

/***************************************************
 * 5. CSV UPLOAD
 ***************************************************/
function processCSVUpload(csvData) {

  try {

    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    let sheet =
      ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.ACTIVATIONS_NAME);
    }

    const csvRows = Utilities.parseCsv(csvData);

    if (csvRows.length < 2) {
      return "CSV EMPTY";
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

    return "SUCCESS";

  } catch (e) {

    return "ERROR : " + e.toString();

  }

}

/***************************************************
 * 6. NEW ACTIVATIONS
 ***************************************************/
function processNewActivations(dataRows) {

  try {

    const ss = SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(SHEETS.ACTIVATIONS_NAME);

    const existingData =
      sheet.getDataRange().getValues();

    const existingIMEIs =
      existingData.map(r =>
        String(r[0]).trim()
      );

    const rowsToAppend = [];

    dataRows.forEach(rowStr => {

      const cols = rowStr.split('\t');

      const imei = String(cols[0]).trim();

      if (
        imei &&
        existingIMEIs.indexOf(imei) === -1
      ) {

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
      status: "duplicate",
      count: 0
    };

  } catch (e) {

    return {
      status: "error",
      message: e.toString()
    };

  }

}

/***************************************************
 * 7. RECENT REGISTRATIONS
 ***************************************************/
function getRecentSubmissions() {

  try {

    const ss =
      SpreadsheetApp.openById(REG_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.REGISTRATIONS_NAME
      );

    if (!sheet) return [];

    const data =
      sheet.getDataRange().getValues();

    if (data.length < 2) return [];

    const headers = data[0];

    return data.slice(1).map(row => {

      let obj = {};

      headers.forEach((header, i) => {

        let val = row[i];

        if (val instanceof Date) {

          val = Utilities.formatDate(
            val,
            TIMEZONE,
            "dd-MM-yyyy"
          );

        }

        if (
          String(header).toUpperCase().includes("URL") ||
          String(header).toUpperCase().includes("FILE")
        ) {

          val = convertToDirectDownload(val);

        }

        obj[header] = val;

      });

      return obj;

    }).reverse();

  } catch (e) {

    return [];

  }

}

/***************************************************
 * 8. REQUIREMENT HISTORY
 ***************************************************/
function getRequirementData() {

  try {

    const ss =
      SpreadsheetApp.openById(REQ_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.REQUIREMENTS_NAME
      );

    if (!sheet) {

      return {
        rows: [],
        totalCount: 0,
        headers: []
      };

    }

    const data =
      sheet.getDataRange().getValues();

    if (data.length < 1) {

      return {
        rows: [],
        totalCount: 0,
        headers: []
      };

    }

    const headers = data[0];

    const rows = data.slice(1).map((row, rowIndex) => {

      let obj = {};

      obj["_sheetRow"] = rowIndex + 2;

      headers.forEach((header, i) => {

        let val = row[i];

        if (val instanceof Date) {

          val = Utilities.formatDate(
            val,
            TIMEZONE,
            "dd-MM-yyyy HH:mm"
          );

        }

        obj[header] =
          (val === "" || val === null)
            ? "-"
            : val;

      });

      return obj;

    }).reverse();

    return {
      rows: rows,
      headers: headers,
      totalCount: rows.length
    };

  } catch (e) {

    return {
      rows: [],
      totalCount: 0,
      error: e.toString()
    };

  }

}

/***************************************************
 * 9. REQUIREMENT SUBMIT
 ***************************************************/
function processRequirement(data) {

  try {

    const ss =
      SpreadsheetApp.openById(REQ_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.REQUIREMENTS_NAME
      );

    if (!sheet) {
      throw new Error("Requirements sheet not found");
    }

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

    return e.toString();

  }

}

/***************************************************
 * 10. DELETE REQUIREMENT
 ***************************************************/
function deleteRequirement(rowNumber) {

  try {

    const ss =
      SpreadsheetApp.openById(REQ_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.REQUIREMENTS_NAME
      );

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    sheet.deleteRow(Number(rowNumber));

    return "SUCCESS";

  } catch (e) {

    return "ERROR : " + e.toString();

  }

}

/***************************************************
 * 11. UTILITIES
 ***************************************************/
function convertToDirectDownload(url) {

  if (
    !url ||
    typeof url !== 'string' ||
    !url.includes("drive.google.com")
  ) {
    return url;
  }

  const match = url.match(/[-\w]{25,}/);

  return match
    ? "https://drive.google.com/uc?export=download&id=" + match[0]
    : url;

}

/***************************************************
 * 12. PERFORMANCE OVERVIEW
 ***************************************************/
function getDashboardStats() {

  try {

    const ss =
      SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.ACTIVATIONS_NAME
      );

    if (!sheet) return defaultStats();

    const data =
      sheet.getDataRange().getValues();

    if (data.length < 2) {
      return defaultStats();
    }

    return {
      totalCount: data.length - 1,
      lastMonthCount: 0,
      growthPercent: 0,
      outletActive: 0,
      totalRetailers: 0,
      trend: []
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
 * 13. UNIQUE RETAILERS
 ***************************************************/
function getUniqueRetailers() {

  try {

    const ss =
      SpreadsheetApp.openById(ACT_SPREADSHEET_ID);

    const sheet =
      ss.getSheetByName(
        SHEETS.ACTIVATIONS_NAME
      );

    if (!sheet) return [];

    const data =
      sheet.getDataRange().getValues();

    if (data.length < 2) return [];

    const headers = data[0];

    const retailerIndex =
      headers.findIndex(h =>
        /retailer/i.test(String(h))
      );

    if (retailerIndex === -1) return [];

    const set = new Set();

    data.slice(1).forEach(r => {

      const retailer =
        String(r[retailerIndex] || '').trim();

      if (retailer) {
        set.add(retailer);
      }

    });

    return Array.from(set).sort();

  } catch (e) {

    return [];

  }

}

/***************************************************
 * 14. PRICE LIST
 ***************************************************/
function getProductsByMode(mode) {

  try {

    const sheetName = PRICE_SHEETS[mode];

    if (!sheetName) return [];

    const ss =
      SpreadsheetApp.openById(
        PRICE_SPREADSHEET_ID
      );

    const sheet =
      ss.getSheetByName(sheetName);

    if (!sheet) return [];

    const data =
      sheet.getDataRange().getValues();

    if (data.length < 2) return [];

    const headers = data[0];

    return data.slice(1).map(row => {

      let obj = {};

      headers.forEach((h, i) => {

        let val = row[i];

        if (typeof val === "number") {
          val = Number(val.toFixed(2));
        }

        obj[h] = val;

      });

      return obj;

    });

  } catch (e) {

    return [];

  }

}

/***************************************************
 * 15. GIFT DASHBOARD
 ***************************************************/
function getDashboardEntries() {

  try {

    const ss =
      SpreadsheetApp.openById(
        GIF_SPREADSHEET_ID
      );

    const sheet =
      ss.getSheetByName(
        SHEETS.GIFT_NAME
      );

    if (!sheet) return [];

    const values =
      sheet.getDataRange().getValues();

    if (values.length < 2) return [];

    return values.slice(1).map(row => ({

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

    }));

  } catch (e) {

    return [];

  }

}

/***************************************************
 * 16. TEST FUNCTION
 ***************************************************/
function testDashboard() {

  const data = getDashboardData();

  Logger.log(JSON.stringify(data));

}
