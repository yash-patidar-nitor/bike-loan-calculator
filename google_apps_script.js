/**
 * Google Apps Script for Hero Fincorp Cases (Dhanteras 2026)
 * 
 * Setup Instructions:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1J9KhmScwrm8zpIr2bBzHxq0egicn70rawZh4MOW5QL8/edit
 * 2. In the top menu, go to Extensions > Apps Script
 * 3. Delete any default code in Code.gs, paste this entire script, and click Save (💾)
 * 4. In the top right, click Deploy > New deployment
 * 5. Select type: "Web app"
 * 6. Set Description: "Bike Loan Calculator Webhook"
 * 7. Set Execute as: "Me"
 * 8. Set Who has access: "Anyone"  <-- IMPORTANT
 * 9. Click "Deploy", authorize permissions when prompted, and copy the "Web app URL" (starts with https://script.google.com/macros/s/.../exec)
 * 10. Paste that URL into the Bike Loan Calculator (Settings in Form Fill modal)
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};
    
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // Initialize headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Token No",
        "Timestamp",
        "Status",
        "Assigned Executive",
        "Customer Name",
        "Mobile Number",
        "Pin Code",
        "Mother Name",
        "Mother DOB",
        "Nominee Name",
        "Nominee Relation",
        "Home Mobile",
        "Reference Name",
        "Reference Mobile",
        "Bike Model",
        "Ex-Showroom Price",
        "RTO (12%)",
        "Insurance",
        "Other Charges",
        "On-Road Price",
        "Down Payment (DP)",
        "File Charge (FC)",
        "General Insurance",
        "Tenure (Mo)",
        "Flat ROI (%)",
        "Monthly EMI",
        "Net Landing Cost",
        "Case Link (Click to Open)"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
      sheet.setFrozenRows(1);
    }

    // Calculate Token Number: TK-1001, TK-1002, etc. based on row index
    var nextRow = sheet.getLastRow() + 1;
    var tokenNo = data.token || ("TK-" + (1000 + nextRow - 1));

    var timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    var status = data.status || "New Case";
    var assignedTo = data.assignedTo || "Unassigned";

    var caseUrl = data.caseUrl || "";
    // Hyperlink formula for clean clicking inside Google Sheets
    var linkFormula = caseUrl ? '=HYPERLINK("' + caseUrl + '", "🔗 Open ' + tokenNo + '")' : "";

    var row = [
      tokenNo,
      timestamp,
      status,
      assignedTo,
      data.customerName || "",
      data.mobile || "",
      data.pincode || "",
      data.motherName || "",
      data.motherDob || "",
      data.nomineeName || "",
      data.nomineeRel || "",
      data.homeMobile || "",
      data.refName || "",
      data.refMobile || "",
      data.bikeModel || "",
      data.exShowroom || 0,
      data.rto || 0,
      data.insuranceAmount || 0,
      data.otherCharges || 0,
      data.onRoadPrice || 0,
      data.downPayment || 0,
      data.fileCharge || 0,
      data.generalInsurance ? "Yes (₹1,000)" : "No",
      data.tenureMonths || 0,
      data.roi || 0,
      data.monthlyEmi || 0,
      data.netLanding || 0,
      linkFormula
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      token: tokenNo,
      row: nextRow
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Hero Fincorp Google Apps Script Webhook is Active! Use POST to submit cases.")
    .setMimeType(ContentService.MimeType.TEXT);
}
