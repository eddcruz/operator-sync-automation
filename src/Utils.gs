/** ---- Utilities ---- */

/**
 * @function normalizeName
 * Normalizes a name to enable case-insensitive, accent-insensitive,
 * and order-insensitive comparisons.
 *
 * - Converts to uppercase
 * - Removes accents and diacritic characters
 * - Splits into words (by whitespace)
 * - Sorts words alphabetically
 * - Rejoins into a single string
 *
 * Example:
 *  Input:  "José Luis Pérez"
 *  Output: "JOSE LUIS PEREZ" → ["JOSE", "LUIS", "PEREZ"] → sorted → "JOSE LUIS PEREZ"
 */
function normalizeName(name) {
  return (name || "")
    .toString()
    .toUpperCase()                                   // Convert to uppercase
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove Unicode accents
    .split(/\s+/)                                    // Split by whitespace
    .sort()                                          // Sort alphabetically
    .join(" ")                                       // Rejoin into a single string
    .trim();
}

/**
 * @function extractEmails
 * Extracts plain email addresses from a raw email header string.
 * Handles both "Name <email>" and plain "email" formats.
 *
 * @param {string} raw - Raw header string (From, To, Cc)
 * @returns {string[]} Array of plain email addresses
 */
function extractEmails(raw) {
  if (!raw) return [];
  return raw
    .split(",")
    .map(s => s.trim())
    .map(s => s.replace(/^.*<([^>]+)>\s*$/, "$1"))
    .filter(Boolean);
}

/**
 * @function cleanExtractedName
 * Cleans a name string extracted via regex from an email subject or body.
 * Removes noise keywords and non-letter characters left by partial matches.
 *
 * @param {string} name - Raw extracted name
 * @returns {string} Cleaned uppercase name
 */
function cleanExtractedName(name) {
  return (name || "")
    .toString()
    .replace(/\b(TERMINATION DATE|TERMINATION TYPE|REASON|SUB-REASON)\b.*$/i, "")
    .replace(/[^\p{L}\s.]/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}