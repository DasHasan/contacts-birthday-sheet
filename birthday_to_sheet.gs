/**
 * Fetches all Google Contacts with birthdays and writes them to a Google Sheet.
 * Run this function from the Apps Script editor.
 */
function syncBirthdaysToSheet() {
  Logger.log("=== syncBirthdaysToSheet started ===");
  const sheet = getOrCreateSheet("Birthdays");

  sheet.clearContents();
  sheet.appendRow(["Name", "Birthday", "Day", "Month", "Year"]);

  const rows = [];
  let nextPageToken = null;
  let pageNumber = 0;
  let totalContacts = 0;

  do {
    pageNumber++;
    const params = {
      resourceName: "people/me",
      // Fetch every field that could possibly hold a "Speichern unter" / File-as value
      personFields: "names,birthdays,fileAses",
      pageSize: 1000,
    };
    if (nextPageToken) params.pageToken = nextPageToken;

    Logger.log(`Fetching page ${pageNumber}...`);
    const response = People.People.Connections.list("people/me", params);
    const connections = response.connections || [];
    totalContacts += connections.length;
    Logger.log(`Page ${pageNumber}: ${connections.length} contacts received`);

    for (const person of connections) {
      const name = getDisplayName(person);
      const birthday = getBirthday(person);

      if (birthday) {
        Logger.log(`  [WITH BIRTHDAY] name="${name}" birthday="${birthday.formatted}"`);
        rows.push([
          name,
          birthday.formatted,
          birthday.day,
          birthday.month,
          birthday.year,
        ]);
      }
    }

    nextPageToken = response.nextPageToken;
  } while (nextPageToken);

  Logger.log(`Total contacts fetched: ${totalContacts}`);
  Logger.log(`Contacts with birthdays: ${rows.length}`);

  if (rows.length === 0) {
    Logger.log("No contacts with birthdays — aborting.");
    SpreadsheetApp.getUi().alert("No contacts with birthdays found.");
    return;
  }

  // Sort by month then day
  rows.sort((a, b) => {
    if (a[3] !== b[3]) return a[3] - b[3]; // month
    return a[2] - b[2]; // day
  });

  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  const header = sheet.getRange(1, 1, 1, 5);
  header.setFontWeight("bold");
  header.setBackground("#4285F4");
  header.setFontColor("#FFFFFF");

  sheet.autoResizeColumns(1, 5);

  Logger.log("=== syncBirthdaysToSheet finished ===");
  SpreadsheetApp.getUi().alert(
    `Done! ${rows.length} contacts with birthdays written to the "Birthdays" sheet.`
  );
}

/**
 * DEBUG FUNCTION — run this first to inspect raw API data.
 * Fetches the first 20 contacts and logs ALL name-related fields so we can
 * identify exactly where "Speichern unter" is stored.
 */
function debugNameFields() {
  Logger.log("=== debugNameFields started ===");

  const response = People.People.Connections.list("people/me", {
    resourceName: "people/me",
    personFields: "names,fileAses",
    pageSize: 20,
  });

  const connections = response.connections || [];
  Logger.log(`Fetched ${connections.length} contacts for inspection`);

  for (const person of connections) {
    const resourceName = person.resourceName || "unknown";
    Logger.log(`\n--- Contact: ${resourceName} ---`);

    // names
    const names = person.names || [];
    Logger.log(`  names (${names.length}):`);
    for (const n of names) {
      Logger.log(`    displayName="${n.displayName}" unstructuredName="${n.unstructuredName}" givenName="${n.givenName}" familyName="${n.familyName}" source.type="${n.metadata?.source?.type}"`);
    }

    // fileAses ("Speichern unter")
    const fileAses = person.fileAses || [];
    Logger.log(`  fileAses (${fileAses.length}):`);
    for (const f of fileAses) {
      Logger.log(`    value="${f.value}" source.type="${f.metadata?.source?.type}"`);
    }
  }

  Logger.log("\n=== debugNameFields finished — open View > Logs to review ===");
}

function getDisplayName(person) {
  // "Speichern unter" is stored in the fileAses field of the People API
  const fileAses = person.fileAses || [];
  Logger.log(`    fileAses count: ${fileAses.length}`);
  if (fileAses.length > 0 && fileAses[0].value) {
    Logger.log(`    -> using fileAs: "${fileAses[0].value}"`);
    return fileAses[0].value;
  }

  const names = person.names || [];
  if (names.length > 0) {
    const displayName = names[0].displayName || "(No Name)";
    Logger.log(`    -> using displayName: "${displayName}"`);
    return displayName;
  }

  Logger.log("    -> no name found, returning (No Name)");
  return "(No Name)";
}

function getBirthday(person) {
  const birthdays = person.birthdays;
  if (!birthdays || birthdays.length === 0) return null;

  const bday = birthdays[0].date;
  if (!bday) {
    Logger.log("    birthday entry found but date is null");
    return null;
  }

  const day = bday.day || null;
  const month = bday.month || null;
  const year = bday.year || null;

  if (!day || !month) {
    Logger.log(`    incomplete birthday skipped (day=${day} month=${month})`);
    return null;
  }

  const monthNames = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  ];

  const formatted = year
    ? `${day} ${monthNames[month - 1]} ${year}`
    : `${day} ${monthNames[month - 1]}`;

  return { formatted, day, month, year: year || "" };
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    Logger.log(`Sheet "${name}" not found — creating it`);
    sheet = ss.insertSheet(name);
  } else {
    Logger.log(`Sheet "${name}" found`);
  }
  return sheet;
}
