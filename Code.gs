const SPREADSHEET_ID = '1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MSP Contract Portal - Approvals')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInboxData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('SUBMISSIONS');
  const logSheet = ss.getSheetByName('ACTION_LOGS');
  
  if (!sheet) return JSON.stringify({ error: 'SUBMISSIONS sheet not found.' });

  // 1. Fetch All Action Logs to build the thread history
  let actionLogs = [];
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    logData.shift(); // Remove headers
    actionLogs = logData.map(row => {
      return {
        timestamp: row[0],
        rfpNo: row[1],
        actorEmail: row[2],
        action: row[3],
        remarks: row[4],
        targetEmail: row[5],
        ccEmail: row[6] || ''
      };
    });
  }

  // 2. Fetch Submissions
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  let submissions = data.map((row, index) => {
    const currentRfpNo = row[5] || '';
    
    // Filter logs that match this specific request
    const requestHistory = actionLogs.filter(log => log.rfpNo === currentRfpNo);

    return {
      id: index + 1,
      timestamp: row[1] || '',             
      requestorEmail: row[2] || '',        
      primaryRecipient: row[3] || '',      
      secondaryRecipient: row[4] || '',    
      rfpNo: currentRfpNo,                 
      dueDate: row[6] || '',               
      year: row[7] || '',                  
      month: row[8] || '',                 
      payorName: row[9] || '',             
      payee: row[10] || '',                
      property: row[11] || '',             
      location: row[12] || '',             
      sector: row[13] || '',               
      serviceKind: row[14] || '',          
      contractNo: row[15] || '',           
      contractAmount: row[16] || '',       
      invoiceNo: row[17] || '',            
      billingPeriod: row[18] || '',        
      soaAmount: row[19] || '',            
      status: row[20] || 'Pending',        
      fileLink: row[21] || '',             
      isRead: false,
      history: requestHistory // Attach the thread history here
    };
  });

  return JSON.stringify(submissions.reverse()); 
}

function processAction(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('ACTION_LOGS');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('ACTION_LOGS');
      logSheet.appendRow(['Timestamp', 'RFP/PEF No.', 'Actor Email', 'Action Taken', 'Remarks/Comments', 'Target Email', 'Cc Email']);
      logSheet.getRange("A1:G1").setFontWeight("bold");
    }

    logSheet.appendRow([
      new Date(),
      payload.rfpNo,
      payload.actorEmail,
      payload.action,
      payload.remarks,
      payload.targetEmail,
      payload.ccEmail
    ]);

    return JSON.stringify({ success: true, message: 'Action successfully recorded in database.' });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}
