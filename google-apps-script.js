// ============================================
// 📊 Google Apps Script — Order Storage in Google Sheets
// ============================================
//
// SETUP INSTRUCTIONS:
// 1. Go to https://script.google.com
// 2. Create a new project
// 3. Paste this code
// 4. Click Deploy → New Deployment
// 5. Select 'Web App'
// 6. Set 'Execute as' = Me
// 7. Set 'Who has access' = Anyone
// 8. Deploy and copy the URL
// 9. Paste the URL in assets/js/whatsapp.js as GOOGLE_SHEET_URL
//
// This script creates a 'PichalOrders' sheet in your Google Spreadsheet
// and receives order data via POST requests from the website.
// ============================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    // Append the order row
    sheet.appendRow([
      data.orderId || '',
      data.date || new Date().toISOString(),
      data.customerName || '',
      data.customerPhone || '',
      data.customerAddress || '',
      data.customerCity || '',
      data.customerPincode || '',
      data.customerLandmark || '',
      data.items || '',
      data.subtotal || 0,
      data.delivery || 0,
      data.total || 0,
      data.status || 'New',
      data.customerNotes || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Order saved' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Pichal Orders API is running!' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PichalOrders');

  if (!sheet) {
    sheet = ss.insertSheet('PichalOrders');
    // Add headers
    sheet.appendRow([
      'Order ID',
      'Date',
      'Customer Name',
      'Phone',
      'Address',
      'City',
      'Pincode',
      'Landmark',
      'Items',
      'Subtotal',
      'Delivery',
      'Total',
      'Status',
      'Notes'
    ]);

    // Format header row
    var headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2D5016');
    headerRange.setFontColor('#FFFFFF');

    // Auto-resize columns
    for (var i = 1; i <= 14; i++) {
      sheet.autoResizeColumn(i);
    }

    // Freeze header row
    sheet.setFrozenRows(1);
  }

  return sheet;
}
