function getDeletableSheets() {
  const EXCLUDED_SHEETS = [TEMPLATE_SHEET, DB_SHEET];
  const today = new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const deletable = [];

  for (let sheet of sheets) {
    const name = sheet.getName();
    if (EXCLUDED_SHEETS.includes(name)) continue;

    // Cherche l'extrême droite du nom : format XX AAAA (semaine et année)
    const match = name.match(/(\d{2})\s?(\d{4})$/);
    if (match) {
      const lastWeek = parseInt(match[1], 10);
      const year = parseInt(match[2], 10);
      const targetDate = getDateOfISOWeek(lastWeek, year);

      if (targetDate < today) {
        deletable.push(name);
      }
    }
  }
  return deletable;
}

function deleteSheetAndCleanup(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const dbSheet = ss.getSheetByName(DB_SHEET);
  if (!sheet || !dbSheet) return;

  // 🔍 Étape 1 — On récupère l’année de fin (la plus à droite dans le nom de la feuille)
  const match = sheetName.match(/(\d{4})$/);
  const endYear = match ? parseInt(match[1], 10) : new Date().getFullYear();

  // 🔍 Étape 2 — On lit la ligne 7 : elle contient les dates de la feuille (format JJ/MM)
  const rawDates = sheet.getRange("7:7").getValues()[0];

  // ⏱️ Étape 3 — On prépare une liste de mois pour détecter un éventuel changement d’année
  const months = rawDates.map(cell => {
    if (typeof cell === "string" && cell.includes("/")) {
      return parseInt(cell.split("/")[1], 10);
    }
    return null;
  }).filter(Boolean);

  // 🔁 Étape 4 — On détecte s’il y a un "retour arrière" sur les mois, signe d’un passage d’année
  const hasYearSwitch = months.some((m, i) => i > 0 && m < months[i - 1]);

  // 🧠 Étape 5 — On reconstitue les dates codées (format AAAAMMJJ) avec les bonnes années
  let currentYear = hasYearSwitch ? endYear - 1 : endYear;
  const codedDates = [];

  for (let i = 3; i < rawDates.length; i++) { // On saute les colonnes A, B, C
    const cell = rawDates[i];
    if (typeof cell === "string" && cell.includes("/")) {
      const [day, month] = cell.split("/").map(n => parseInt(n, 10));

      // Si on détecte un passage de décembre à janvier, on change d’année
      if (hasYearSwitch && i > 3) {
        const prevMonth = parseInt(rawDates[i - 1].split("/")[1], 10);
        if (month < prevMonth) currentYear++;
      }

      const code = `${currentYear}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
      codedDates.push(code);
    }
  }

  // 🧹 Étape 6 — On filtre la DB pour ne garder que les lignes qui NE SONT PAS dans les dates à supprimer
  const dbData = dbSheet.getDataRange().getValues();
  const header = dbData[0]; // ligne d'en-tête
  const body = dbData.slice(1); // le reste

  const filtered = body.filter(row => !codedDates.includes(String(row[0])));

  // 🧾 Étape 7 — On nettoie et réinjecte la DB
  dbSheet.clearContents();
  if (filtered.length > 0) {
    dbSheet.getRange(1, 1, filtered.length + 1, header.length).setValues([header, ...filtered]);
  } else {
    dbSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // 🗑️ Étape 8 — Enfin, on supprime la feuille
  ss.deleteSheet(sheet);
}