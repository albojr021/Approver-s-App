const DB_MAIN_ID = '1YAvZmCdWXbjOcJA-uUY40e6qVqzyiHcB06NpiPcz6y4';
const DB_USER_ID = '1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI';
const UPLOAD_FOLDER_ID = '1ctjUaEFZPe7YLGu1GlFB7BPwWPQeB0SK';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal Approvers Workspace')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let hexStr = '';
  for (let i = 0; i < digest.length; i++) {
    let byte = digest[i];
    if (byte < 0) byte += 256;
    let byteStr = byte.toString(16);
    if (byteStr.length == 1) byteStr = '0' + byteStr;
    hexStr += byteStr;
  }
  return hexStr;
}

function requestOTP(email, intent) {
  try {
    const ss = SpreadsheetApp.openById(DB_USER_ID);
    const sheet = ss.getSheetByName('PEP');
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      let emailExists = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][2].toString().toLowerCase() === email.toLowerCase()) {
          emailExists = true;
          break;
        }
      }
      
      if (intent === 'register' && emailExists) {
        return JSON.stringify({ success: false, message: 'Email is already registered. Please log in.' });
      }
      if (intent === 'reset' && !emailExists) {
        return JSON.stringify({ success: false, message: 'Email not found in our records.' });
      }
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const cache = CacheService.getScriptCache();
    cache.put(email + '_' + intent, otp, 600); 
    
    const subject = "Portal Security - Your Verification Code";
    const body = "Your 6-digit verification code is: " + otp + "\n\nThis code will expire in 10 minutes. Do not share this with anyone.";
    
    MailApp.sendEmail(email, subject, body);
    return JSON.stringify({ success: true, message: 'OTP sent to ' + email });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}

function verifyOTP(email, intent, submittedOtp) {
  const cache = CacheService.getScriptCache();
  const cachedOtp = cache.get(email + '_' + intent);
  
  if (!cachedOtp) return JSON.stringify({ success: false, message: 'OTP expired or invalid.' });
  if (cachedOtp === submittedOtp) {
    cache.remove(email + '_' + intent);
    return JSON.stringify({ success: true });
  } else {
    return JSON.stringify({ success: false, message: 'Incorrect OTP.' });
  }
}

function registerUser(payload) {
  try {
    const ss = SpreadsheetApp.openById(DB_USER_ID);
    const sheet = ss.getSheetByName('PEP');
    if (!sheet) return JSON.stringify({ success: false, message: 'PEP sheet not found.' });

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === payload.email) {
        return JSON.stringify({ success: false, message: 'Email is already registered.' });
      }
    }

    const isOrganic = payload.email.endsWith('@megaworld-lifestyle.com') ? 'Organic' : 'Non-organic';
    const hashedPassword = hashPassword(payload.password);
    
    const usernames = sheet.getRange("C:C").getValues();
    let targetRow = sheet.getLastRow() + 1; 
    for (let i = 1; i < usernames.length; i++) { 
      if (usernames[i][0] === "") {
        targetRow = i + 1; 
        break;
      }
    }

    const rowData = [
      new Date(), payload.fullName, payload.email, hashedPassword, 
      'APPROVER', isOrganic, 'Pending', '', ''
    ];
    
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    return JSON.stringify({ success: true, message: 'Registration submitted. Please wait for admin approval.' });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}

function loginUser(email, password) {
  try {
    const ss = SpreadsheetApp.openById(DB_USER_ID);
    const sheet = ss.getSheetByName('PEP');
    const data = sheet.getDataRange().getValues();
    
    const hashedAttempt = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) { 
        
        const dbPassword = data[i][3].toString();
        if (dbPassword !== hashedAttempt && dbPassword !== password) {
          return JSON.stringify({ success: false, message: 'Invalid credentials.' });
        }
        
        if (data[i][7] === 'RESIGNED') { 
          return JSON.stringify({ success: false, message: 'Account deactivated. Status: RESIGNED.' });
        }
        if (data[i][6] !== 'Approved') { 
          return JSON.stringify({ success: false, message: 'Account is still pending approval or restricted.' });
        }
        
        return JSON.stringify({ 
          success: true, 
          user: { fullName: data[i][1], email: data[i][2], title: data[i][4] } 
        });
      }
    }
    return JSON.stringify({ success: false, message: 'Account not found.' });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}

function resetPassword(email, newPassword) {
  try {
    const ss = SpreadsheetApp.openById(DB_USER_ID);
    const sheet = ss.getSheetByName('PEP');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) {
        sheet.getRange(i + 1, 4).setValue(hashPassword(newPassword));
        return JSON.stringify({ success: true, message: 'Password updated successfully.' });
      }
    }
    return JSON.stringify({ success: false, message: 'Account not found.' });
  } catch(error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}

function checkUserStatus(email) {
  const ss = SpreadsheetApp.openById(DB_USER_ID);
  const sheet = ss.getSheetByName('PEP');
  if(!sheet) return JSON.stringify({ valid: false });
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      if (data[i][7] === 'RESIGNED' || data[i][6] !== 'Approved') {
        return JSON.stringify({ valid: false });
      }
      return JSON.stringify({ valid: true });
    }
  }
  return JSON.stringify({ valid: false });
}

function getInboxData(userEmail) {
  if(!userEmail) return JSON.stringify({ error: 'Unauthorized access.' });

  const ss = SpreadsheetApp.openById(DB_MAIN_ID);
  const sheet = ss.getSheetByName('SUBMISSIONS');
  const logSheet = ss.getSheetByName('ACTION_LOGS');
  
  if (!sheet) return JSON.stringify({ error: 'SUBMISSIONS sheet not found.' });

  let actionLogs = [];
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    logData.shift(); 
    actionLogs = logData.map(row => {
      return { timestamp: row[0], rfpNo: row[1], actorEmail: row[2], action: row[3], remarks: row[4], targetEmail: row[5], ccEmail: row[6] || '', fileLink: row[7] || '' };
    });
  }

  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  let submissions = [];
  
  data.forEach((row, index) => {
    const primaryTo = row[3] || ''; 
    const ccListStr = row[4] || ''; 
    const requestor = row[2] || ''; 
    const currentRfpNo = row[5] || '';

    const queryEmail = userEmail.trim().toLowerCase();
    
    // Original recipients validation
    const toArray = primaryTo.toLowerCase().split(',').map(e => e.trim());
    const ccArray = ccListStr.toLowerCase().split(',').map(e => e.trim());

    // Routing History validation
    const requestHistory = actionLogs.filter(log => log.rfpNo === currentRfpNo);
    let isInHistory = false;
    
    // Check if the current user was ever part of the routing (target, cc, or actor)
    for (let i = 0; i < requestHistory.length; i++) {
      let log = requestHistory[i];
      if (log.actorEmail.toLowerCase() === queryEmail || 
          log.targetEmail.toLowerCase().includes(queryEmail) || 
          log.ccEmail.toLowerCase().includes(queryEmail)) {
          isInHistory = true;
          break;
      }
    }

    if (toArray.includes(queryEmail) || ccArray.includes(queryEmail) || requestor.toLowerCase() === queryEmail || isInHistory) {
      
      submissions.push({
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
        history: requestHistory 
      });
    }
  });

  return JSON.stringify(submissions.reverse()); 
}

function processAction(payload) {
  try {
    const ss = SpreadsheetApp.openById(DB_MAIN_ID);
    let logSheet = ss.getSheetByName('ACTION_LOGS');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('ACTION_LOGS');
      logSheet.appendRow(['Timestamp', 'RFP/PEF No.', 'Actor Email', 'Action Taken', 'Remarks/Comments', 'Target Email', 'Cc Email', 'File Upload']);
      logSheet.getRange("A1:H1").setFontWeight("bold");
    }

    let fileUrl = '';
    if (payload.fileObj && payload.fileObj.fileData) {
      const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
      const decodedData = Utilities.base64Decode(payload.fileObj.fileData);
      const blob = Utilities.newBlob(decodedData, payload.fileObj.mimeType, payload.fileObj.fileName);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }

    logSheet.appendRow([ new Date(), payload.rfpNo, payload.actorEmail, payload.action, payload.remarks, payload.targetEmail, payload.ccEmail, fileUrl ]);
    return JSON.stringify({ success: true, fileLink: fileUrl });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString() });
  }
}

function processPendingEmails() {
  const ss = SpreadsheetApp.openById(DB_USER_ID);
  const sheet = ss.getSheetByName('PEP');
  if(!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let permStatus = data[i][6]; 
    let emailSent = data[i][8];  
    let userEmail = data[i][2];  
    let fullName = data[i][1];   
    
    if (permStatus === 'Approved' && emailSent !== 'Done' && userEmail !== '') {
      try {
        let subject = "Portal Account Approved";
        let body = "Hello " + fullName + ",\n\nYour account has been approved by the administrator. You may now log in to the Portal system.\n\nThank you.";
        MailApp.sendEmail(userEmail, subject, body);
        
        sheet.getRange(i + 1, 9).setValue('Done');
      } catch(e) {
         Logger.log("Failed to send email to: " + userEmail);
      }
    }
  }
}
