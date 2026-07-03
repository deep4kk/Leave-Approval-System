/**
 * Leave Application & Approval System
 * Google Apps Script Web Application
 * 
 * SETUP INSTRUCTIONS:
 * Create the following sheets in your Google Sheet:
 * 
 * 1. "Users" - Columns: Email, Name, Role, Department, Active
 *    Roles: Employee, Manager, Admin
 * 
 * 2. "LeaveRequests" - Columns: RequestID, EmployeeEmail, EmployeeName, LeaveType, 
 *    StartDate, EndDate, WorkingDays, Reason, Status, ManagerEmail, 
 *    SubmittedAt, ApprovedAt, RejectedAt, RejectionReason
 * 
 * 3. "LeaveBalances" - Columns: EmployeeEmail, LeaveType, TotalDays, UsedDays, RemainingDays, Year
 * 
 * 4. "AuditLog" - Columns: Timestamp, User, Action, RecordID, OldValue, NewValue
 * 
 * Deploy as Web App with:
 * - Execute as: Me
 * - Access: Anyone with Google account
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  SHEET_NAMES: {
    USERS: 'Users',
    LEAVE_REQUESTS: 'LeaveRequests',
    LEAVE_BALANCES: 'LeaveBalances',
    AUDIT_LOG: 'AuditLog'
  },
  LEAVE_TYPES: ['Sick', 'Casual', 'Earned'],
  STATUS: {
    PENDING: 'Pending',
    APPROVED: 'Approved',
    REJECTED: 'Rejected',
    CANCELLED: 'Cancelled'
  },
  ROLES: {
    EMPLOYEE: 'Employee',
    MANAGER: 'Manager',
    ADMIN: 'Admin'
  }
};

// ============== AUTHENTICATION ==============
function getUserRole(email) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][4] === true) {
      return data[i][2];
    }
  }
  return null;
}

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  return user || { email, name: 'Unknown', role: null, department: null };
}

function getUserByEmail(email) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) {
      return {
        email: data[i][0],
        name: data[i][1],
        role: data[i][2],
        department: data[i][3],
        active: data[i][4]
      };
    }
  }
  return null;
}

function requireRole(allowedRoles) {
  const user = getCurrentUser();
  if (!user.role || !allowedRoles.includes(user.role)) {
    throw new Error('Unauthorized: Insufficient permissions');
  }
  return user;
}

// ============== SHEET HELPERS ==============
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    setupHeaders(sheet, name);
  }
  return sheet;
}

function setupHeaders(sheet, name) {
  const headers = {
    'Users': ['Email', 'Name', 'Role', 'Department', 'Active'],
    'LeaveRequests': ['RequestID', 'EmployeeEmail', 'EmployeeName', 'LeaveType', 
                      'StartDate', 'EndDate', 'WorkingDays', 'Reason', 'Status', 
                      'ManagerEmail', 'SubmittedAt', 'ApprovedAt', 'RejectedAt', 'RejectionReason'],
    'LeaveBalances': ['EmployeeEmail', 'LeaveType', 'TotalDays', 'UsedDays', 'RemainingDays', 'Year'],
    'AuditLog': ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue']
  };
  
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
  }
}

function generateId() {
  return Utilities.getUuid();
}

// ============== AUDIT LOG ==============
function logAction(action, recordId, oldValue, newValue) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.AUDIT_LOG);
  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail();
  
  sheet.appendRow([timestamp, user, action, recordId, oldValue, newValue]);
}

// ============== NOTIFICATIONS ==============
function sendNotification(toEmail, subject, htmlBody) {
  try {
    MailApp.sendEmail({
      to: toEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Leave Management System'
    });
    return true;
  } catch (e) {
    console.error('Email send failed:', e);
    return false;
  }
}

function getEmailTemplate(status, employeeName, leaveType, startDate, endDate, reason, managerName) {
  const statusColors = {
    Pending: '#F59E0B',
    Approved: '#10B981',
    Rejected: '#EF4444'
  };
  
  const color = statusColors[status] || '#6B7280';
  
  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; border-radius: 16px 16px 0 0;">
        <h1 style="color: white; margin: 0; font-size: 24px;">Leave Request ${status}</h1>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 16px 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1);">
        <p style="color: #374151; font-size: 16px;">Dear ${employeeName},</p>
        <p style="color: #6B7280; font-size: 14px;">Your leave request has been <strong style="color: ${color};">${status}</strong></p>
        
        <div style="background: #F9FAFB; padding: 20px; border-radius: 12px; margin: 20px 0;">
          <p style="margin: 0 0 10px;"><strong style="color: #374151;">Leave Type:</strong> <span style="color: #6B7280;">${leaveType}</span></p>
          <p style="margin: 0 0 10px;"><strong style="color: #374151;">Duration:</strong> <span style="color: #6B7280;">${formatDate(startDate)} - ${formatDate(endDate)}</span></p>
          <p style="margin: 0;"><strong style="color: #374151;">Reason:</strong> <span style="color: #6B7280;">${reason}</span></p>
        </div>
        
        <div style="text-align: center; margin-top: 30px;">
          <a href="${ScriptApp.getService().getUrl()}" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 30px; border-radius: 25px; text-decoration: none; font-weight: 600;">View in System</a>
        </div>
      </div>
    </div>
  `;
}

function formatDate(date) {
  return new Date(date).toLocaleDateString('en-US', { 
    year: 'numeric', month: 'long', day: 'numeric' 
  });
}

// ============== BUSINESS LOGIC ==============
function calculateWorkingDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  let count = 0;
  
  const current = new Date(start);
  while (current <= end) {
    const dayOfWeek = current.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      count++;
    }
    current.setDate(current.getDate() + 1);
  }
  
  return count;
}

function checkLeaveBalance(email, leaveType, daysRequested) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_BALANCES);
  const data = sheet.getDataRange().getValues();
  const year = new Date().getFullYear();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && 
        data[i][1] === leaveType && 
        data[i][5] == year) {
      return {
        available: data[i][4] >= daysRequested,
        remaining: data[i][4],
        requested: daysRequested
      };
    }
  }
  return { available: false, remaining: 0, requested: daysRequested };
}

function checkConflicts(email, startDate, endDate, excludeRequestId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const conflicts = [];
  const user = getUserByEmail(email);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] !== 'Approved') continue;
    if (excludeRequestId && data[i][0] === excludeRequestId) continue;
    
    const otherEmail = data[i][1];
    const otherUser = getUserByEmail(otherEmail);
    if (otherUser.department !== user.department) continue;
    
    const reqStart = new Date(data[i][4]);
    const reqEnd = new Date(data[i][5]);
    const checkStart = new Date(startDate);
    const checkEnd = new Date(endDate);
    
    if (checkStart <= reqEnd && checkEnd >= reqStart) {
      conflicts.push({
        employee: data[i][2],
        dates: `${formatDate(data[i][4])} - ${formatDate(data[i][5])}`
      });
    }
  }
  
  return conflicts;
}

function getManagerForEmployee(email) {
  const user = getUserByEmail(email);
  if (!user) return null;
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === CONFIG.ROLES.MANAGER && data[i][3] === user.department) {
      return data[i][0];
    }
  }
  return null;
}

// ============== API FUNCTIONS ==============
function submitLeaveRequest(leaveType, startDate, endDate, reason) {
  const user = requireRole([CONFIG.ROLES.EMPLOYEE, CONFIG.ROLES.MANAGER, CONFIG.ROLES.ADMIN]);
  
  const workingDays = calculateWorkingDays(startDate, endDate);
  const balance = checkLeaveBalance(user.email, leaveType, workingDays);
  
  if (!balance.available) {
    throw new Error(`Insufficient ${leaveType} leave balance. Available: ${balance.remaining} days, Requested: ${workingDays} days`);
  }
  
  const conflicts = checkConflicts(user.email, startDate, endDate);
  if (conflicts.length > 0) {
    throw new Error(`Conflict detected with approved leave from ${conflicts[0].employee} (${conflicts[0].dates})`);
  }
  
  const requestId = generateId();
  const managerEmail = getManagerForEmployee(user.email);
  const now = new Date();
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  sheet.appendRow([
    requestId, user.email, user.name, leaveType, 
    new Date(startDate), new Date(endDate), workingDays, reason,
    CONFIG.STATUS.PENDING, managerEmail, now, '', '', ''
  ]);
  
  logAction('SUBMIT_LEAVE_REQUEST', requestId, '', CONFIG.STATUS.PENDING);
  
  // Send notification to manager
  if (managerEmail) {
    const approvalToken = generateApprovalToken(requestId, 'approve');
    const rejectToken = generateApprovalToken(requestId, 'reject');
    const approveUrl = `${ScriptApp.getService().getUrl()}?action=approve&token=${approvalToken}`;
    const rejectUrl = `${ScriptApp.getService().getUrl()}?action=reject&token=${rejectToken}`;
    
    sendManagerNotification(managerEmail, user.name, leaveType, startDate, endDate, reason, approveUrl, rejectUrl);
  }
  
  return { success: true, requestId };
}

function generateApprovalToken(requestId, action) {
  const payload = JSON.stringify({ requestId, action, timestamp: Date.now() });
  const base64 = Utilities.base64Encode(payload);
  return base64.replace(/=+$/, '');
}

function parseApprovalToken(token) {
  try {
    const padded = token + '===';
    const decoded = Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString();
    return JSON.parse(decoded);
  } catch (e) {
    return null;
  }
}

function sendManagerNotification(managerEmail, employeeName, leaveType, startDate, endDate, reason, approveUrl, rejectUrl) {
  const htmlBody = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #10B981 0%, #059669 100%); padding: 30px; border-radius: 16px 16px 0 0;">
        <h1 style="color: white; margin: 0;">Leave Request Approval Required</h1>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 16px 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1);">
        <p style="color: #374151; font-size: 16px;">Hi Manager,</p>
        <p style="color: #6B7280;"><strong>${employeeName}</strong> has submitted a leave request that requires your approval.</p>
        
        <div style="background: #F9FAFB; padding: 20px; border-radius: 12px; margin: 20px 0;">
          <p style="margin: 0 0 10px;"><strong>Type:</strong> ${leaveType}</p>
          <p style="margin: 0 0 10px;"><strong>Duration:</strong> ${formatDate(startDate)} - ${formatDate(endDate)}</p>
          <p style="margin: 0;"><strong>Reason:</strong> ${reason}</p>
        </div>
        
        <div style="text-align: center; margin-top: 30px;">
          <a href="${approveUrl}" style="background: #10B981; color: white; padding: 12px 30px; border-radius: 25px; text-decoration: none; font-weight: 600; margin-right: 10px;">Approve</a>
          <a href="${rejectUrl}" style="background: #EF4444; color: white; padding: 12px 30px; border-radius: 25px; text-decoration: none; font-weight: 600;">Reject</a>
        </div>
      </div>
    </div>
  `;
  
  sendNotification(managerEmail, `Leave Request from ${employeeName}`, htmlBody);
}

function approveLeaveRequest(requestId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.APPROVED);
      sheet.getRange(row, 11, 1, 1).setValue(new Date());
      
      // Update leave balance
      const email = data[i][1];
      const leaveType = data[i][3];
      const workingDays = data[i][6];
      updateLeaveBalance(email, leaveType, workingDays);
      
      logAction('APPROVE_LEAVE', requestId, oldStatus, CONFIG.STATUS.APPROVED);
      
      // Notify employee
      const employeeEmail = data[i][1];
      const employeeName = data[i][2];
      sendNotification(employeeEmail, 'Leave Request Approved', 
        getEmailTemplate(CONFIG.STATUS.APPROVED, employeeName, leaveType, data[i][4], data[i][5], data[i][7], 'Manager'));
      
      return { success: true };
    }
  }
  
  throw new Error('Request not found');
}

function rejectLeaveRequest(requestId, reason) {
  if (!reason) throw new Error('Rejection reason is required');
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.REJECTED);
      sheet.getRange(row, 12, 1, 1).setValue(new Date());
      sheet.getRange(row, 13, 1, 1).setValue(reason);
      
      logAction('REJECT_LEAVE', requestId, oldStatus, CONFIG.STATUS.REJECTED);
      
      // Notify employee
      const employeeEmail = data[i][1];
      const employeeName = data[i][2];
      const leaveType = data[i][3];
      sendNotification(employeeEmail, 'Leave Request Rejected', 
        getEmailTemplate(CONFIG.STATUS.REJECTED, employeeName, leaveType, data[i][4], data[i][5], reason, 'Manager'));
      
      return { success: true };
    }
  }
  
  throw new Error('Request not found');
}

function updateLeaveBalance(email, leaveType, daysUsed) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_BALANCES);
  const data = sheet.getDataRange().getValues();
  const year = new Date().getFullYear();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && 
        data[i][1] === leaveType && 
        data[i][5] == year) {
      const row = i + 1;
      const used = data[i][3] + daysUsed;
      const remaining = data[i][4] - daysUsed;
      
      sheet.getRange(row, 4, 1, 1).setValue(used);
      sheet.getRange(row, 5, 1, 1).setValue(remaining);
      return;
    }
  }
}

function cancelLeaveRequest(requestId) {
  requireRole([CONFIG.ROLES.EMPLOYEE, CONFIG.ROLES.ADMIN]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      const user = getCurrentUser();
      if (data[i][1] !== user.email && user.role !== CONFIG.ROLES.ADMIN) {
        throw new Error('You can only cancel your own requests');
      }
      
      const row = i + 1;
      const oldStatus = data[i][8];
      
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.CANCELLED);
      logAction('CANCEL_LEAVE', requestId, oldStatus, CONFIG.STATUS.CANCELLED);
      
      return { success: true };
    }
  }
  
  throw new Error('Request not found');
}

// ============== DATA RETRIEVAL ==============
function getMyLeaveRequests() {
  const user = getCurrentUser();
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toLowerCase() === user.email.toLowerCase()) {
      requests.push({
        requestId: data[i][0],
        leaveType: data[i][3],
        startDate: data[i][4],
        endDate: data[i][5],
        workingDays: data[i][6],
        reason: data[i][7],
        status: data[i][8],
        submittedAt: data[i][10]
      });
    }
  }
  
  return requests;
}

function getMyBalances() {
  const user = getCurrentUser();
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_BALANCES);
  const data = sheet.getDataRange().getValues();
  const year = new Date().getFullYear();
  const balances = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === user.email.toLowerCase() && data[i][5] == year) {
      balances.push({
        leaveType: data[i][1],
        total: data[i][2],
        used: data[i][3],
        remaining: data[i][4]
      });
    }
  }
  
  return balances;
}

function getPendingApprovals() {
  const user = getCurrentUser();
  if (user.role !== CONFIG.ROLES.MANAGER && user.role !== CONFIG.ROLES.ADMIN) {
    throw new Error('Only managers can view pending approvals');
  }
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === CONFIG.STATUS.PENDING) {
      if (user.role === CONFIG.ROLES.ADMIN || data[i][9] === user.email) {
        requests.push({
          requestId: data[i][0],
          employeeEmail: data[i][1],
          employeeName: data[i][2],
          leaveType: data[i][3],
          startDate: data[i][4],
          endDate: data[i][5],
          workingDays: data[i][6],
          reason: data[i][7],
          submittedAt: data[i][10]
        });
      }
    }
  }
  
  return requests;
}

function getAllRequests() {
  requireRole([CONFIG.ROLES.ADMIN]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.LEAVE_REQUESTS);
  const data = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    requests.push({
      requestId: data[i][0],
      employeeEmail: data[i][1],
      employeeName: data[i][2],
      leaveType: data[i][3],
      startDate: data[i][4],
      endDate: data[i][5],
      workingDays: data[i][6],
      reason: data[i][7],
      status: data[i][8],
      managerEmail: data[i][9],
      submittedAt: data[i][10],
      approvedAt: data[i][11],
      rejectedAt: data[i][12],
      rejectionReason: data[i][13]
    });
  }
  
  return requests;
}

// ============== WEB APP ==============
function doGet(e) {
  const htmlOutput = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Leave Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // Handle token-based actions from email links
  if (e.parameter.action && e.parameter.token) {
    const token = parseApprovalToken(e.parameter.token);
    if (token && token.requestId) {
      try {
        if (e.parameter.action === 'approve') {
          approveLeaveRequest(token.requestId);
          return HtmlService.createHtmlOutput('<script>alert("Leave request approved!");google.script.host.close();</script>');
        } else if (e.parameter.action === 'reject') {
          return HtmlService.createHtmlOutput(`
            <script>
              function submitReject() {
                const reason = document.getElementById('reason').value;
                if (!reason) {
                  alert('Please provide a reason');
                  return;
                }
                google.script.run
                  .withSuccessHandler(() => {
                    alert('Leave request rejected');
                    google.script.host.close();
                  })
                  .withFailureHandler(e => alert(e.message))
                  .rejectLeaveRequest('${token.requestId}', reason);
              }
            </script>
            <div style="font-family: sans-serif; padding: 40px; text-align: center;">
              <h2>Reject Leave Request</h2>
              <p>Please provide a reason for rejection:</p>
              <textarea id="reason" rows="4" style="width: 100%; margin: 10px 0;"></textarea>
              <button onclick="submitReject()" style="padding: 10px 30px; background: #EF4444; color: white; border: none; border-radius: 5px; cursor: pointer;">Reject</button>
            </div>
          `);
        }
      } catch (e) {
        return HtmlService.createHtmlOutput(`<script>alert("Error: ${e.message}");google.script.host.close();</script>`);
      }
    }
  }
  
  return htmlOutput;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Export for testing
const test = {
  calculateWorkingDays,
  checkLeaveBalance,
  generateApprovalToken,
  parseApprovalToken
};
