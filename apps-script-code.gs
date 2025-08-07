// Google Apps Script for Salon Client Consultation & Consumption Forms
// Web App URL: https://script.google.com/macros/s/AKfycbynE1KRwABodBVmrzC3jep5LweAP0B53pUPW89PgL_EGrx8IKH5ikMj2EN3QWY6a2ZThQ/exec

function doPost(e) {
  try {
    // Parse the incoming data
    const formData = JSON.parse(e.parameter.formData);
    const formType = formData.formType;
    const submitId = formData.submitId;
    
    // Log the submission for debugging
    console.log(`Received ${formType} form submission with ID: ${submitId}`);
    
    // Process based on form type
    if (formType === 'consultation') {
      return processConsultationForm(formData);
    } else if (formType === 'consumption') {
      return processConsumptionForm(formData);
    } else {
      throw new Error('Invalid form type specified');
    }
    
  } catch (error) {
    console.error('Error processing form submission:', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processConsultationForm(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('wpConsultation');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = createConsultationSheet(spreadsheet);
    } else {
      // Check if headers exist, if not add them
      const firstRow = sheet.getRange(1, 1, 1, 18).getValues()[0];
      if (!firstRow[0] || firstRow[0] === 'A' || firstRow[0] === '') {
        // Headers don't exist or are default, add them
        addConsultationHeaders(sheet);
      }
    }
    
    // Save files to Drive and get links
    const signatureLink = data.clientSignature ? saveFileToDrive(data.clientSignature, `signature_${data.submitId}`, 'png') : '';
    const beforePhotoLink = data.beforePhoto ? saveFileToDrive(data.beforePhoto, `before_photo_${data.submitId}`, 'jpg') : '';
    
    // Prepare row data for consultation
    const rowData = [
      new Date(), // Timestamp
      data.submitId || '',
      data.clientName || '',
      data.phone || '',
      data.email || '',
      data.visitDate || '',
      data.clientType || '',
      data.services ? data.services.join(', ') : '',
      data.hairType || '',
      data.hairTexture || '',
      data.scalpCondition || '',
      data.previousTreatments || '',
      data.allergies || '',
      data.clientExpectations || '',
      data.consent || '',
      signatureLink, // Digital Signature Drive Link
      beforePhotoLink, // Before Photo Drive Link
      'Pending' // Status
    ];
    
    // Append data to sheet
    sheet.appendRow(rowData);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 18);
    
    console.log(`Consultation form saved successfully with ID: ${data.submitId}`);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Consultation form saved successfully',
        submitId: data.submitId
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error saving consultation form:', error);
    throw error;
  }
}

function processConsumptionForm(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('wpConsumption');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = createConsumptionSheet(spreadsheet);
    } else {
      // Check if headers exist, if not add them
      const firstRow = sheet.getRange(1, 1, 1, 15).getValues()[0];
      if (!firstRow[0] || firstRow[0] === 'A' || firstRow[0] === '') {
        // Headers don't exist or are default, add them
        addConsumptionHeaders(sheet);
      }
    }
    
    const timestamp = new Date();
    const submitId = data.submitId || '';
    const consultationId = data.consultationId || '';
    const totalCost = data.totalCost || '';
    const serviceNotes = data.serviceNotes || '';
    const clientSatisfaction = data.clientSatisfaction || '';
    const followUp = data.followUp || '';
    
    // Save after photo to Drive and get link
    const afterPhotoLink = data.afterPhoto ? saveFileToDrive(data.afterPhoto, `after_photo_${data.submitId}`, 'jpg') : '';
    
    // Get arrays of products and stylists
    const productNames = data.productName || [];
    const productTypes = data.productType || [];
    const quantitiesUsed = data.quantityUsed || [];
    const units = data.unit || [];
    const stylistNames = data.stylistName || [];
    const incentiveSplits = data.incentiveSplit || [];
    
    // Create separate rows for each product
    const maxProducts = Math.max(productNames.length, productTypes.length, quantitiesUsed.length, units.length);
    for (let i = 0; i < maxProducts; i++) {
      const rowData = [
        timestamp, // Timestamp
        submitId, // Submit ID
        consultationId, // Consultation ID
        productNames[i] || '', // Product Name
        productTypes[i] || '', // Product Type
        quantitiesUsed[i] || '', // Quantity Used
        units[i] || '', // Unit
        '', // Stylist Name (will be filled in stylist rows)
        '', // Incentive Split (will be filled in stylist rows)
        i === 0 ? totalCost : '', // Total Cost (only in first row)
        i === 0 ? serviceNotes : '', // Service Notes (only in first row)
        i === 0 ? clientSatisfaction : '', // Client Satisfaction (only in first row)
        i === 0 ? followUp : '', // Follow-up (only in first row)
        i === 0 ? afterPhotoLink : '', // After Photo (only in first row)
        i === 0 ? 'Completed' : '' // Status (only in first row)
      ];
      
      sheet.appendRow(rowData);
    }
    
    // Create separate rows for each stylist
    const maxStylists = Math.max(stylistNames.length, incentiveSplits.length);
    for (let i = 0; i < maxStylists; i++) {
      const rowData = [
        timestamp, // Timestamp
        submitId, // Submit ID
        consultationId, // Consultation ID
        '', // Product Name
        '', // Product Type
        '', // Quantity Used
        '', // Unit
        stylistNames[i] || '', // Stylist Name
        incentiveSplits[i] || '', // Incentive Split
        '', // Total Cost
        '', // Service Notes
        '', // Client Satisfaction
        '', // Follow-up
        '', // After Photo
        '' // Status
      ];
      
      sheet.appendRow(rowData);
    }
    
    // Update consultation sheet with consumption data
    updateConsultationWithConsumption(data);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 15);
    
    console.log(`Consumption form saved successfully with ID: ${data.submitId}`);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Consumption form saved successfully',
        submitId: data.submitId
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error saving consumption form:', error);
    throw error;
  }
}

function createConsultationSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('wpConsultation');
  
  // Set up headers
  const headers = [
    'Timestamp',
    'Submit ID',
    'Client Name',
    'Phone Number',
    'Email Address',
    'Visit Date',
    'Client Type',
    'Desired Services',
    'Hair Type',
    'Hair Texture',
    'Scalp Condition',
    'Previous Treatments',
    'Allergies/Sensitivities',
    'Client Expectations',
    'Consent Given',
    'Digital Signature',
    'Before Photo',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#8B5CF6');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Submit ID
  sheet.setColumnWidth(3, 150); // Client Name
  sheet.setColumnWidth(4, 120); // Phone
  sheet.setColumnWidth(5, 180); // Email
  sheet.setColumnWidth(6, 100); // Visit Date
  sheet.setColumnWidth(7, 100); // Client Type
  sheet.setColumnWidth(8, 200); // Services
  sheet.setColumnWidth(9, 100); // Hair Type
  sheet.setColumnWidth(10, 100); // Hair Texture
  sheet.setColumnWidth(11, 120); // Scalp Condition
  sheet.setColumnWidth(12, 200); // Previous Treatments
  sheet.setColumnWidth(13, 200); // Allergies
  sheet.setColumnWidth(14, 250); // Expectations
  sheet.setColumnWidth(15, 100); // Consent
  sheet.setColumnWidth(16, 100); // Signature
  sheet.setColumnWidth(17, 100); // Before Photo
  sheet.setColumnWidth(18, 100); // Status
  
  return sheet;
}

function createConsumptionSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('wpConsumption');
  
  // Set up headers
  const headers = [
    'Timestamp',
    'Submit ID',
    'Consultation ID',
    'Product Names',
    'Product Types',
    'Quantities Used',
    'Units',
    'Stylists/Technicians',
    'Incentive Splits',
    'Total Cost',
    'Service Notes',
    'Client Satisfaction',
    'Follow-up Recommendations',
    'After Photo',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#06B6D4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Submit ID
  sheet.setColumnWidth(3, 120); // Consultation ID
  sheet.setColumnWidth(4, 200); // Product Names
  sheet.setColumnWidth(5, 150); // Product Types
  sheet.setColumnWidth(6, 120); // Quantities
  sheet.setColumnWidth(7, 80);  // Units
  sheet.setColumnWidth(8, 200); // Stylists
  sheet.setColumnWidth(9, 150); // Incentive Splits
  sheet.setColumnWidth(10, 100); // Total Cost
  sheet.setColumnWidth(11, 300); // Service Notes
  sheet.setColumnWidth(12, 150); // Satisfaction
  sheet.setColumnWidth(13, 250); // Follow-up
  sheet.setColumnWidth(14, 100); // After Photo
  sheet.setColumnWidth(15, 100); // Status
  
  return sheet;
}

function addConsultationHeaders(sheet) {
  // Set up headers
  const headers = [
    'Timestamp',
    'Submit ID',
    'Client Name',
    'Phone Number',
    'Email Address',
    'Visit Date',
    'Client Type',
    'Desired Services',
    'Hair Type',
    'Hair Texture',
    'Scalp Condition',
    'Previous Treatments',
    'Allergies/Sensitivities',
    'Client Expectations',
    'Consent Given',
    'Digital Signature',
    'Before Photo',
    'Status'
  ];
  
  // Insert headers at the top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#8B5CF6');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Submit ID
  sheet.setColumnWidth(3, 150); // Client Name
  sheet.setColumnWidth(4, 120); // Phone
  sheet.setColumnWidth(5, 180); // Email
  sheet.setColumnWidth(6, 100); // Visit Date
  sheet.setColumnWidth(7, 100); // Client Type
  sheet.setColumnWidth(8, 200); // Services
  sheet.setColumnWidth(9, 100); // Hair Type
  sheet.setColumnWidth(10, 100); // Hair Texture
  sheet.setColumnWidth(11, 120); // Scalp Condition
  sheet.setColumnWidth(12, 200); // Previous Treatments
  sheet.setColumnWidth(13, 200); // Allergies
  sheet.setColumnWidth(14, 250); // Expectations
  sheet.setColumnWidth(15, 100); // Consent
  sheet.setColumnWidth(16, 100); // Signature
  sheet.setColumnWidth(17, 100); // Before Photo
  sheet.setColumnWidth(18, 100); // Status
}

function addConsumptionHeaders(sheet) {
  // Set up headers
  const headers = [
    'Timestamp',
    'Submit ID',
    'Consultation ID',
    'Product Names',
    'Product Types',
    'Quantities Used',
    'Units',
    'Stylists/Technicians',
    'Incentive Splits',
    'Total Cost',
    'Service Notes',
    'Client Satisfaction',
    'Follow-up Recommendations',
    'After Photo',
    'Status'
  ];
  
  // Insert headers at the top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#06B6D4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Submit ID
  sheet.setColumnWidth(3, 120); // Consultation ID
  sheet.setColumnWidth(4, 200); // Product Names
  sheet.setColumnWidth(5, 150); // Product Types
  sheet.setColumnWidth(6, 120); // Quantities
  sheet.setColumnWidth(7, 80);  // Units
  sheet.setColumnWidth(8, 200); // Stylists
  sheet.setColumnWidth(9, 150); // Incentive Splits
  sheet.setColumnWidth(10, 100); // Total Cost
  sheet.setColumnWidth(11, 300); // Service Notes
  sheet.setColumnWidth(12, 150); // Satisfaction
  sheet.setColumnWidth(13, 250); // Follow-up
  sheet.setColumnWidth(14, 100); // After Photo
  sheet.setColumnWidth(15, 100); // Status
}

function updateConsultationWithConsumption(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const consultationSheet = spreadsheet.getSheetByName('wpConsultation');
    
    if (!consultationSheet) {
      console.log('Consultation sheet not found, skipping update');
      return;
    }
    
    // Find the consultation record by consultation ID
    const consultationId = data.consultationId;
    if (!consultationId) {
      console.log('No consultation ID provided, skipping update');
      return;
    }
    
    // Search for the consultation ID in column B (Submit ID) or column C (Client Name)
    const dataRange = consultationSheet.getDataRange();
    const values = dataRange.getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === consultationId || values[i][2] === consultationId) {
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (rowIndex === -1) {
      console.log(`Consultation ID ${consultationId} not found in consultation sheet`);
      return;
    }
    
    // Update the consultation record status to Completed
    consultationSheet.getRange(rowIndex, 18, 1, 1).setValue('Completed');
    
    console.log(`Updated consultation record at row ${rowIndex} status to Completed`);
    
  } catch (error) {
    console.error('Error updating consultation with consumption data:', error);
  }
}

function saveFileToDrive(base64Data, fileName, fileType) {
  try {
    // Remove data URL prefix if present
    let base64String = base64Data;
    if (base64Data.includes(',')) {
      base64String = base64Data.split(',')[1];
    }
    
    // Decode base64 to bytes
    const bytes = Utilities.base64Decode(base64String);
    
    // Create blob
    const blob = Utilities.newBlob(bytes, `image/${fileType}`, fileName);
    
    // Get the folder where the spreadsheet is located
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
    
    // Create a subfolder for salon files if it doesn't exist
    let salonFolder;
    try {
      salonFolder = folder.getFoldersByName('Salon Files').next();
    } catch (e) {
      // Create the folder if it doesn't exist
      salonFolder = folder.createFolder('Salon Files');
    }
    
    // Save file to Drive
    const file = salonFolder.createFile(blob);
    
    // Set file permissions to anyone with link can view
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    console.log(`File saved to Drive: ${file.getName()} - ${file.getUrl()}`);
    
    return file.getUrl();
    
  } catch (error) {
    console.error('Error saving file to Drive:', error);
    return 'Error saving file';
  }
}

function doGet(e) {
  // Handle GET requests (for testing or data retrieval)
  return ContentService
    .createTextOutput(JSON.stringify({
      message: 'Salon Forms Web App is running',
      timestamp: new Date().toISOString(),
      endpoints: {
        consultation: 'POST with formType: consultation',
        consumption: 'POST with formType: consumption'
      }
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Utility function to test the web app
function testWebApp() {
  const testData = {
    formType: 'consultation',
    submitId: 'TEST001',
    clientName: 'Test Client',
    phone: '1234567890',
    email: 'test@example.com',
    visitDate: '2024-01-15',
    clientType: 'new',
    services: ['haircut', 'coloring'],
    hairType: 'straight',
    hairTexture: 'medium',
    scalpCondition: 'normal',
    previousTreatments: 'None',
    allergies: 'None',
    clientExpectations: 'Test expectations',
    consent: 'yes',
    clientSignature: 'data:image/png;base64,test',
    beforePhoto: 'data:image/jpeg;base64,test'
  };
  
  const payload = {
    formData: JSON.stringify(testData)
  };
  
  const options = {
    method: 'POST',
    payload: payload
  };
  
  try {
    const response = UrlFetchApp.fetch(ScriptApp.getService().getUrl(), options);
    console.log('Test response:', response.getContentText());
  } catch (error) {
    console.error('Test failed:', error);
  }
}

// Function to set up the spreadsheet with proper formatting
function setupSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheets if they don't exist
  if (!spreadsheet.getSheetByName('wpConsultation')) {
    createConsultationSheet(spreadsheet);
  }
  
  if (!spreadsheet.getSheetByName('wpConsumption')) {
    createConsumptionSheet(spreadsheet);
  }
  
  // Set up data validation and formatting
  setupDataValidation(spreadsheet);
  
  console.log('Spreadsheet setup completed successfully');
}

function setupDataValidation(spreadsheet) {
  const consultationSheet = spreadsheet.getSheetByName('wpConsultation');
  const consumptionSheet = spreadsheet.getSheetByName('wpConsumption');
  
  if (consultationSheet) {
    // Add data validation for Client Type
    const clientTypeRange = consultationSheet.getRange('G2:G1000');
    const clientTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['new', 'returning', 'walk-in'], true)
      .setAllowInvalid(false)
      .setHelpText('Please select a valid client type')
      .build();
    clientTypeRange.setDataValidation(clientTypeRule);
    
    // Add data validation for Status
    const statusRange = consultationSheet.getRange('R2:R1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Completed', 'Cancelled'], true)
      .setAllowInvalid(false)
      .setHelpText('Please select a valid status')
      .build();
    statusRange.setDataValidation(statusRule);
  }
  
  if (consumptionSheet) {
    // Add data validation for Status
    const statusRange = consumptionSheet.getRange('O2:O1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Completed', 'Cancelled'], true)
      .setAllowInvalid(false)
      .setHelpText('Please select a valid status')
      .build();
    statusRange.setDataValidation(statusRule);
  }
}

// Function to clean up old test data (optional)
function cleanupTestData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const consultationSheet = spreadsheet.getSheetByName('wpConsultation');
  const consumptionSheet = spreadsheet.getSheetByName('wpConsumption');
  
  if (consultationSheet) {
    const dataRange = consultationSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Remove rows with test data (Submit ID starting with 'TEST')
    for (let i = values.length - 1; i > 0; i--) {
      if (values[i][1] && values[i][1].toString().startsWith('TEST')) {
        consultationSheet.deleteRow(i + 1);
      }
    }
  }
  
  if (consumptionSheet) {
    const dataRange = consumptionSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Remove rows with test data (Submit ID starting with 'TEST')
    for (let i = values.length - 1; i > 0; i--) {
      if (values[i][1] && values[i][1].toString().startsWith('TEST')) {
        consumptionSheet.deleteRow(i + 1);
      }
    }
  }
  
  console.log('Test data cleanup completed');
}

// Function to fix existing sheets that don't have proper headers
function fixExistingSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Fix consultation sheet
  const consultationSheet = spreadsheet.getSheetByName('wpConsultation');
  if (consultationSheet) {
    const firstRow = consultationSheet.getRange(1, 1, 1, 18).getValues()[0];
    if (!firstRow[0] || firstRow[0] === 'A' || firstRow[0] === '') {
      console.log('Fixing consultation sheet headers...');
      addConsultationHeaders(consultationSheet);
    }
  }
  
  // Fix consumption sheet
  const consumptionSheet = spreadsheet.getSheetByName('wpConsumption');
  if (consumptionSheet) {
    const firstRow = consumptionSheet.getRange(1, 1, 1, 15).getValues()[0];
    if (!firstRow[0] || firstRow[0] === 'A' || firstRow[0] === '') {
      console.log('Fixing consumption sheet headers...');
      addConsumptionHeaders(consumptionSheet);
    }
  }
  
  console.log('Existing sheets fixed successfully');
}
