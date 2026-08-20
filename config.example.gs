/**
 * config.example.gs
 *
 * Copy this file, rename it to config.gs, and fill in your values.
 * Never commit config.gs to version control.
 *
 * How to use:
 * 1. In Apps Script, create a new file called config.gs
 * 2. Paste this content and replace all placeholder values
 * 3. Save and run onOpen() to verify the menu appears
 */

// ── Email settings ──────────────────────────────────────────────
// Email address that sends the weekly operator roster (Excel attachment)
const HR_SENDER_EMAIL = "hr.sender@yourcompany.com";

// Subject keyword used to identify roster emails
const HR_EMAIL_SUBJECT = "Weekly active staff list";

// Email addresses that may send operator deduction queries
const DEDUCTION_SENDERS = [
  "hr1@yourcompany.com",
  "hr2@yourcompany.com",
  "operations@yourcompany.com"
];

// Your email — receives alerts and sends auto-replies
const ALERTS_EMAIL = "you@yourcompany.com";

// ── Google Drive ─────────────────────────────────────────────────
// ID of the Drive folder where incoming Excel files are stored
// Get it from the folder URL: drive.google.com/drive/folders/THIS_PART
const DRIVE_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID";

// ── Google Sheets ─────────────────────────────────────────────────
// Sheet tab name for the operator catalog
const OPERATORS_SHEET_NAME = "operators_catalog";

// Sheet tab name for deduction records
const DEDUCTIONS_SHEET_NAME = "operator_missing_items";

// ── Source sheet names inside the HR Excel files ──────────────────
// These are the tab names inside the Excel attachments from HR
const HR_EXCEL_SHEETS = ["Company Division A", "Company Division B"];

// ── Safety threshold ──────────────────────────────────────────────
// Minimum number of active IDs required to trigger status reconciliation.
// Prevents accidental mass-deactivation if HR sends an incomplete file.
const MIN_ACTIVE_IDS_THRESHOLD = 20;

// ── Gmail label ───────────────────────────────────────────────────
// Label applied to processed threads to avoid reprocessing
const PROCESSED_LABEL = "Sync_Operators_Processed";
const DEDUCTIONS_LABEL = "Sync_Deductions_Processed";