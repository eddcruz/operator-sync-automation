/************************************************************
 * Operator ingestion and reconciliation (HR → Drive → Sheets)
 * - Inserts new operators (no duplicates by Operator_ID)
 * - Status reconciliation: Active/Inactive
 * - Safety threshold + alert if the list is suspiciously small
 *
 * Requires: Drive API (Advanced Service) enabled.
 ************************************************************/

// === Configurable parameters ===
const MIN_ACTIVE_IDS_THRESHOLD = 20; // minimum IDs seen to trigger reconciliation
const ALERTS_EMAIL = "you@yourcompany.com";

// === Main flow: process emails, import and reconcile ===
function procesarCorreosOperadores() {
    const sender       = "hr.sender@yourcompany.com";
    const subjectKey   = "Weekly active staff list";
    const driveFolderId = "YOUR_DRIVE_FOLDER_ID";
    const sheetId      = SpreadsheetApp.getActiveSpreadsheet().getId();
    const sheetName    = "operators_catalog";
    let newOperators   = 0;

    const threads = GmailApp.search(
        `from:(${sender}) subject:(${subjectKey}) newer_than:90d -label:Sync_Operators_HR_Drive_Gmail`
    );
    const folder  = DriveApp.getFolderById(driveFolderId);
    const label   = GmailApp.getUserLabelByName("Sync_Operators_HR_Drive_Gmail") ||
        GmailApp.createLabel("Sync_Operators_HR_Drive_Gmail");

    // Active IDs seen across all attachments in this run
    const activeIdsThisRun = new Set();

    threads.forEach(thread => {
        thread.getMessages().forEach(msg => {
            msg.getAttachments().forEach(file => {
                if (file.getContentType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
                    const fileName  = file.getName();
                    const alreadyExists = folder.getFilesByName(fileName).hasNext();

                    if (!alreadyExists) {
                        // Save and process
                        const savedFile = folder.createFile(file);
                        const r = importFromExcel(savedFile.getBlob(), sheetId, sheetName);
                        newOperators += r.newCount;
                        r.activeIds.forEach(id => activeIdsThisRun.add(id));
                    } else {
                        // (Optional) collect IDs for reconciliation even if we skip saving the duplicate
                        const r = importFromExcel(file.copyBlob(), sheetId, sheetName, /*collectOnly*/ true);
                        r.activeIds.forEach(id => activeIdsThisRun.add(id));
                    }
                }
            });
        });
        label.addToThread(thread);
    });

    // Notify if new operators were added
    if (newOperators > 0) {
        MailApp.sendEmail({
            to: ALERTS_EMAIL,
            subject: "✔️ New operators added",
            body: `${newOperators} new operators have been added to the operators_catalog sheet.`
        });
    }

    // Status reconciliation with safety threshold
    const totalSeen = activeIdsThisRun.size;
    if (totalSeen >= MIN_ACTIVE_IDS_THRESHOLD) {
        reconcileOperatorStatus(Array.from(activeIdsThisRun));
    } else if (totalSeen > 0) {
        // Small list: alert and do NOT reconcile to protect the database
        MailApp.sendEmail({
            to: ALERTS_EMAIL,
            subject: "⚠️ Operator list suspiciously small: reconciliation NOT applied",
            htmlBody: `<p>A list with only <strong>${totalSeen}</strong> active IDs was processed (threshold: ${MIN_ACTIVE_IDS_THRESHOLD}).</p>
                 <p>The <strong>Status</strong> column was not updated to avoid marking false terminations.</p>
                 <p>Please review the HR attachment and re-run when the full list is available.</p>`
        });
    }
}

/**
 * Imports data from an Excel file:
 *  - Converts to Google Sheets (temporary)
 *  - Inserts new records (deduped by Operator_ID)
 *  - Collects and returns active IDs seen
 *  - Deletes the temporary file after extraction
 *
 * @param {Blob}    blob         - Excel file blob
 * @param {string}  sheetId      - Destination spreadsheet ID
 * @param {string}  sheetName    - Destination sheet tab name
 * @param {boolean} collectOnly  - If true, do NOT insert rows; only read IDs
 * @return {{ newCount: number, activeIds: string[] }}
 */
function importFromExcel(blob, sheetId, sheetName, collectOnly) {
    const converted = Drive.Files.insert({ title: blob.getName(), mimeType: MimeType.GOOGLE_SHEETS }, blob);
    const file      = SpreadsheetApp.openById(converted.id);

    const sourceTabs = ["Division A", "Division B"]; // Tab names inside the HR Excel file
    const destSheet  = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

    // Load existing IDs to detect duplicates before inserting
    const lastRow    = destSheet.getLastRow();
    const existingIds = lastRow > 1
        ? destSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(id => id.toString().trim())
        : [];

    let newCount = 0;
    const activeIdsFound = new Set();

    sourceTabs.forEach(tabName => {
        const sheet = file.getSheetByName(tabName);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        if (!data || data.length < 2) return;

        const headers = data[0];

        // Locate columns by header name — update these to match your Excel file
        const idCol   = headers.indexOf("EmployeeId");
        const nameCol = headers.indexOf("EmployeeName");
        const posCol  = headers.indexOf("PositionDescription");
        const roleCol = headers.indexOf("RoleAttribute");

        for (let i = 1; i < data.length; i++) {
            const id = (data[i][idCol] ?? "").toString().trim();
            if (!id) continue;

            // Collect all IDs seen (for status reconciliation)
            activeIdsFound.add(id);

            // Insert new rows only if not in collect-only mode
            if (!collectOnly && !existingIds.includes(id)) {
                destSheet.appendRow([
                    id,
                    data[i][nameCol],
                    data[i][posCol],
                    data[i][roleCol]
                ]);
                existingIds.push(id);
                newCount++;
            }
        }
    });

    // Delete temporary file to keep Drive clean
    DriveApp.getFileById(converted.id).setTrashed(true);

    return { newCount, activeIds: Array.from(activeIdsFound) };
}

/**
 * Reconciles Status in 'operators_catalog':
 *  - "Active"   if Operator_ID ∈ activeIds
 *  - "Inactive" if Operator_ID ∉ activeIds
 * Batch write for performance.
 *
 * @param {string[]} activeIds
 */
function reconcileOperatorStatus(activeIds) {
    if (!activeIds || !activeIds.length) return;

    const ss   = SpreadsheetApp.getActive();
    const sh   = ss.getSheetByName("operators_catalog");
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;

    const headers  = data[0];
    const idCol    = headers.indexOf("Operator_ID");
    const statusCol = headers.indexOf("Status");
    if (idCol === -1 || statusCol === -1) {
        throw new Error("Columns 'Operator_ID' and/or 'Status' not found in operators_catalog.");
    }

    const activeSet    = new Set(activeIds.map(v => v.toString().trim()));
    const body         = data.slice(1);

    const updatedStatus = body.map(row => {
        const id = (row[idCol] ?? "").toString().trim();
        return activeSet.has(id) ? "Active" : "Inactive";
    });

    sh.getRange(2, statusCol + 1, updatedStatus.length, 1)
        .setValues(updatedStatus.map(v => [v]));
}

/* ==========================================================
   Second workflow: automated replies to deduction queries
   ========================================================== */

function revisarFaltantesOperador() {
    const hrSenders = [
        "hr1@yourcompany.com",
        "hr2@yourcompany.com",
        "hr3@yourcompany.com",
        "hr4@yourcompany.com",
        "hr5@yourcompany.com"
        // Add more senders as needed
    ];

    const sheet         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("operator_missing_items");
    const processedLabel = GmailApp.getUserLabelByName("Sync_Operators_HR_Deductions_Processed") ||
        GmailApp.createLabel("Sync_Operators_HR_Deductions_Processed");

    const searchQuery =
        `(${hrSenders.map(r => "from:" + r).join(" OR ")}) ` +
        `(subject:(ACCOUNT OR OPERATOR OR EMPLOYEE OR TERMINATION OR DEDUCTION)) ` +
        `newer_than:3d -label:Sync_Operators_HR_Deductions_Processed`;

    const threads = GmailApp.search(searchQuery);
    if (!threads.length) return;

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return;

    const headers         = data[0];
    const rows            = data.slice(1);

    const colName         = headers.indexOf("Operator_Name");
    const colMovement     = headers.indexOf("Movement_Type");
    const colConcept      = headers.indexOf("Deduction_Concept");
    const colAmount       = headers.indexOf("Total_Amount");
    const colInstallments = headers.indexOf("Installments");

    threads.forEach(thread => {
        const alreadyProcessed = thread.getLabels().some(
            label => label.getName() === "Sync_Operators_HR_Deductions_Processed"
        );
        if (alreadyProcessed) return;

        const messages   = thread.getMessages();
        if (!messages.length) return;

        const firstMsg = messages[0];
        const lastMsg  = messages[messages.length - 1];

        const operatorName = getOperatorName(firstMsg, lastMsg);
        if (!operatorName) return;

        const normalizedSearch = normalizeName(operatorName);
        const matches = rows.filter(r => normalizeName(r[colName] || "") === normalizedSearch);

        let htmlBody = `<p>Hello,</p>`;

        if (matches.length === 0) {
            htmlBody += `<p>Operator <strong>${operatorName}</strong> has no recorded equipment deductions.</p><p>Best regards</p>`;
        } else {
            let summary = {}, grandTotal = 0, grandInstallments = 0;

            matches.forEach(r => {
                const key = `${r[colMovement]} - ${r[colConcept]}`;
                if (!summary[key]) summary[key] = { amount: 0, installments: 0 };
                const amt  = parseFloat((r[colAmount] || "0").toString().replace(/[^\d.-]/g, "")) || 0;
                const inst = parseInt(r[colInstallments] || 0, 10) || 0;
                summary[key].amount       += amt;
                summary[key].installments += inst;
            });

            htmlBody += `<p>Please find the requested information.</p>`;
            htmlBody += `<p><strong>Equipment Deduction Summary</strong><br>(Operator: <strong>${operatorName}</strong>)</p>`;
            htmlBody += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
                    <tr><th>Movement - Concept</th><th>Total Amount</th><th># Installments</th></tr>`;

            Object.keys(summary).forEach(key => {
                const { amount, installments } = summary[key];
                grandTotal        += amount;
                grandInstallments += installments;
                htmlBody += `<tr><td>${key}</td><td>$${amount.toFixed(2)}</td><td>${installments}</td></tr>`;
            });

            htmlBody += `<tr style="font-weight: bold;"><td>Grand Total:</td><td>$${grandTotal.toFixed(2)}</td><td>${grandInstallments}</td></tr></table><p>Best regards</p>`;
        }

        const myEmail = (Session.getActiveUser().getEmail() || "").toLowerCase();

        // Restore CC recipients from the original message
        const origFrom = extractEmails(firstMsg.getFrom());
        const origTo   = extractEmails(firstMsg.getTo());
        const origCc   = extractEmails(firstMsg.getCc());
        const shouldInclude = new Set([...origFrom, ...origTo, ...origCc].map(e => e.toLowerCase()));

        const lastFrom = extractEmails(lastMsg.getFrom());
        const lastTo   = extractEmails(lastMsg.getTo());
        const lastCc   = extractEmails(lastMsg.getCc());
        const alreadyIncluded = new Set([...lastFrom, ...lastTo, ...lastCc].map(e => e.toLowerCase()));

        const missingCc = [...shouldInclude].filter(e => e && !alreadyIncluded.has(e) && e !== myEmail);

        lastMsg.replyAll("", {
            htmlBody: htmlBody,
            cc: missingCc.join(",")
        });

        processedLabel.addToThread(thread);
    });
}

/**
 * @function getOperatorName
 * Extracts an operator name from the email subject or body
 * using a cascade of regex patterns. Returns null if not found.
 *
 * @param {GmailMessage} firstMsg
 * @param {GmailMessage} lastMsg
 * @returns {string|null}
 */
function getOperatorName(firstMsg, lastMsg) {
    const subjectRaw = (lastMsg.getSubject()    || "").trim();
    const subject    = subjectRaw.toUpperCase();
    const bodyRaw    = (firstMsg.getPlainBody() || "").trim();
    const body       = bodyRaw.toUpperCase();

    // 1) Subject: "OPERATOR: NAME" or "EMPLOYEE: NAME"
    let m =
        subject.match(/OPERAT[O0]R?\s*[\/\-:]\s*([A-ZÁÉÍÓÚÑ.\s]+)$/i) ||
        subject.match(/EMPLOYEE\s*[\/\-:]\s*([A-ZÁÉÍÓÚÑ.\s]+)$/i);
    if (m && m[1]) return cleanExtractedName(m[1]);

    // 2) Subject: "TERMINATION ... - NAME"
    m = subjectRaw.match(/-\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.\s]{6,})\s*$/i);
    if (m && m[1]) return cleanExtractedName(m[1]);

    // 3) Body: table/field format
    m = body.match(
        /NAME\s+(.+?)(?=\s+TERMINATION\s+DATE|\s+TERMINATION\s+TYPE|\s+REASON\b|\s+SUB-REASON\b|$)/is
    );
    if (m && m[1]) return cleanExtractedName(m[1]);

    // 4) Body: "OPERATOR - NAME"
    m = body.match(/OPERATOR[^\n]*\n[-–]\s*([A-ZÁÉÍÓÚÑ.\s]+)\b/mi);
    if (!m) m = body.match(/OPERATOR[^\-]*[-–]\s*([A-ZÁÉÍÓÚÑ.\s]+)\b/mi);
    if (m && m[1]) return cleanExtractedName(m[1]);

    // 5) Body: "for Operator NAME"
    m = body.match(/OPERATOR\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.\s]{8,})(?:,|\n|$)/i);
    if (m && m[1]) return cleanExtractedName(m[1]);

    return null;
}