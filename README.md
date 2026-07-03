# Leave Application & Approval System

A complete Google Apps Script Web Application for managing employee leave requests with multi-level approval workflow.

## Features

- **Role-Based Access Control**: Employee (Requester), Manager (Approver), HR Admin
- **Leave Balance Management**: Auto-calculate working days, validate against remaining balance
- **Conflict Detection**: Flag overlapping leave requests within the same department
- **Email Approvals**: Approve/Reject directly from email without logging in
- **Audit Logging**: Track all state changes with timestamps
- **Apple-Style UI**: Modern, responsive interface with smooth animations

## Setup

1. Create a new Google Apps Script project
2. Create the following sheets:
   - **Users**: Email, Name, Role, Department, Active
   - **LeaveRequests**: RequestID, EmployeeEmail, LeaveType, StartDate, EndDate, WorkingDays, Reason, Status, ManagerEmail, SubmittedAt, ApprovedAt, RejectedAt, RejectionReason
   - **LeaveBalances**: EmployeeEmail, LeaveType, TotalDays, UsedDays, RemainingDays, Year
   - **AuditLog**: Timestamp, User, Action, RecordID, OldValue, NewValue

3. Deploy as Web App (Execute as: Me, Access: Anyone with Google account)

## Architecture

- Google OAuth via Session.getActiveUser()
- State Machine: Draft → Pending → Approved/Rejected → Cancelled
- Token-based email approval links
- Server-side validation on all operations

## Files

- `Code.gs` - Backend server functions
- `Index.html` - Main UI template
- `Stylesheet.html` - Apple-style CSS
- `JavaScript.html` - Client-side logic
