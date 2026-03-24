function triggerFullSyncForUser(sheetName) {
  const { successCount, failCount } = synchronizeFromSheet(sheetName);
  const remaining = countRemainingCalendarSyncsForUser();
  return { successCount, failCount, remaining };
}

function synchronizeFromSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const dbSheet = ss.getSheetByName(DB_SHEET);
  if (!sheet || !dbSheet) return;

  const data = sheet.getDataRange().getValues();
  const dateRow = data[LINE_DATE - 1];
  const baseYear = getBaseYearFromName(sheetName);

  let successCount = 0;
  let failCount = 0;
  const dbData = dbSheet.getDataRange().getValues();

  for (let r = LINE_DATE; r < data.length; r++) {
    const name       = data[r][1];
    const calendarRaw= data[r][2];
    if (!name || !calendarRaw) continue;

    for (let c = 3; c < data[0].length; c++) {
      const value    = data[r][c];
      const dateCell = dateRow[c];
      if (!isValidDateCell(dateCell)) continue;

      const dateCode = buildDateCode(dateCell, baseYear);
      const existing = dbData.find(row =>
        row[COL_DATECODE] == dateCode && row[COL_NAME] == name);

      // Si l'ancienne valeur et la nouvelle sont identiques ET déjà synchronisé → rien à faire
      if (existing
          && existing[COL_OLDVALUE] == value
          && existing[COL_STATUS]   == "Synchronisé") {
        continue;
      }
      // Si pas de nouvelle valeur et pas d'ancienne ligne → on ignore
      if (!value && !existing) continue;

      // Sinon on insère ou met à jour
      const result = updateOrInsertDBRow(dateCode, name, calendarRaw, value, dbSheet);
      result === 1 ? successCount++ : failCount++;
    }
  }

  return { successCount, failCount };
}

// helper pour l’année de base
function getBaseYearFromName(name) {
  const m = name.match(/\d{2}\s?(\d{4})/);
  return m ? parseInt(m[1],10) : new Date().getFullYear();
}
function testGetBaseYearFromName() {
  const aws = getBaseYearFromName("S49 2025 - S04 2026");
  console.log(aws);
}
function isValidDateCell(s) {
  return typeof s === 'string' && s.includes('/');
}
function buildDateCode(cell, baseYear) {
  const [day,mon] = cell.split('/').map(Number);
  let year = baseYear;
  // si on affiche janvier alors qu’on est en novembre/décembre, on incrémente
  if (mon===1 && new Date().getMonth()>=10) year++;
  return `${year}${String(mon).padStart(2,'0')}${String(day).padStart(2,'0')}`;
}

function updateOrInsertDBRow(dateCode, name, calendarRaw, newValue, dbSheet) {
  newValue = newValue || "";
  const data = dbSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[COL_DATECODE] == dateCode && r[COL_NAME] == name);
  
  try {
    if (rowIndex > -1) {
      const oldInDB = data[rowIndex][COL_OLDVALUE];
      const status = data[rowIndex][COL_STATUS];
      
      if (newValue == oldInDB && status == 'Non synchronisé') {
        if (newValue == '' || oldInDB == '') {
          // Cas : oldValue == newValue est vide avant que la sync ne se fasse
          dbSheet.deleteRow(rowIndex + 1);
          // SpreadsheetApp.getActiveSpreadsheet().toast(`🗑️ Événement ${newValue} supprimé pour ${name}`);
          return 1;
        } else {
          // Cas : oldValue == newValue a supprimé une valeur avant que la sync ne se fasse
          dbSheet.getRange(rowIndex + 1, COL_STATUS + 1).setValue("Synchronisé");
          // SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Événement ${newValue} synchronisé pour ${name}`);
          return 1;
        }
      } else {
        // Mise à jour d'une ligne existante
        dbSheet.getRange(rowIndex + 1, COL_NEWVALUE + 1).setValue(newValue);
        dbSheet.getRange(rowIndex + 1, COL_STATUS + 1).setValue("Non synchronisé");
        // SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Événement ${newValue} mise à jour pour ${name}`);
        return 1;
      }
    } else {
      // Nouvelle ligne
      dbSheet.appendRow([
        dateCode,
        name,
        calendarRaw,
        "",              // OLDVALUE vide au départ
        newValue,
        "Non synchronisé"
      ]);
      // SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Événement ${newValue} ajouté pour ${name}`);
      return 1;
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`⚠️ Erreur update DB: ${err.message}`);
    return 0;
  }
}

function showToast(msg) {
  SpreadsheetApp.getActive().toast(msg);
}


function getSyncableSheets() {
  const EXCLUDED = [TEMPLATE_SHEET, DB_SHEET];
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  return sheets
    .map(s => s.getName())
    .filter(name => !EXCLUDED.includes(name));
}

/**
 * Renvoie l'email de l'utilisateur connecté.
 */
function getActiveUserEmail() {
  try {
    return Session
      .getEffectiveUser()
      .getEmail() || "Utilisateur inconnu";
  } catch (e) {
    return "Utilisateur inconnu";
  }
}

/**
 * Compte les événements en attente de sync **pour l'utilisateur courant**.
 */
function countRemainingCalendarSyncsForUser() {
  const email = Session.getEffectiveUser().getEmail();
  const local = email.split('@')[0];       // partie « prenom_nom »
  const db    = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(DB_SHEET);
  if (!db) return 0;

  const data = db.getDataRange().getValues();
  const COL = { NAME:1, STATUS:5 };

  // on parcourt toutes les lignes en filtre “Non synchronisé” & nom correspond
  return data.slice(1).reduce((count, row) => {
    if (String(row[COL.STATUS]).trim() === 'Non synchronisé') {
      // normalise le nom B pour comparaison
      const nomNorm = sanitizeName(row[COL.NAME]);
      if (nomNorm === local) count++;
    }
    return count;
  }, 0);
}