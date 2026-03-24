function generatePlanningBetweenYearWeeks(start, end) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(TEMPLATE_SHEET);

  if (!template) throw new Error(`Feuille modèle "${TEMPLATE_SHEET}" introuvable.`);

  const [startYear, startWeek] = start.split('-').map(Number);
  const [endYear, endWeek] = end.split('-').map(Number);

  const weekDates = getWeekDateRanges(startYear, startWeek, endYear, endWeek);

  // Créer un nom de feuille : SS AAAA - SS AAAA
  const newSheetName = `S${startWeek.toString().padStart(2, '0')} ${startYear} - S${endWeek.toString().padStart(2, '0')} ${endYear}`;
  const newSheet = template.copyTo(ss).setName(newSheetName);
  ss.setActiveSheet(newSheet);

  // Obtenir la cellule de départ
  const startRange = newSheet.getRange(START_CELL_DATE);
  const row = startRange.getRow();
  const col = startRange.getColumn();

  let dates = [];
  weekDates.forEach(({ year, week }) => {
    const monday = getDateOfISOWeek(week, year);
    for (let i = 0; i < (MAX_WEEKS - 1); i++) {
      const date = new Date(monday);
      date.setDate(date.getDate() + i);
      date.setHours(0, 0, 0, 0);
      dates.push([Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'dd/MM')]);
    }
  });

  // Injecter les dates horizontalement depuis START_CELL_DATE
  const flatDates = dates.map(d => d[0]);
  newSheet.getRange(row, col, 1, flatDates.length).setValues([flatDates]);

  // Injecter les SXX à partir de START_CELL_WEEK
  const weekStartRange = newSheet.getRange(START_CELL_WEEK);
  const weekRow = weekStartRange.getRow();
  const weekCol = weekStartRange.getColumn();

  let weekLabels = [];
  weekDates.forEach(({ week }, index) => {
    const label = `S${week.toString().padStart(2, '0')}`;
    weekLabels.push({ label, col: weekCol + index * 7 });
  });

  weekLabels.forEach(({ label, col }) => {
    newSheet.getRange(weekRow, col).setValue(label);
  });
}

/**
 * Retourne un tableau d'objets { year, week } représentant chaque semaine ISO entre deux points.
 * Exemple : de 2025-S24 à 2025-S32 => [ {year: 2025, week: 24}, ..., {year: 2025, week: 32} ]
 */
function getWeekDateRanges(startYear, startWeek, endYear, endWeek) {
  let result = [];
  let currentYear = startYear;
  let currentWeek = startWeek;

  while (currentYear < endYear || (currentYear === endYear && currentWeek <= endWeek)) {
    result.push({ year: currentYear, week: currentWeek });
    currentWeek++;
    if (currentWeek > 52) {
      currentWeek = 1;
      currentYear++;
    }
  }

  return result;
}

/**
 * Retourne la date du lundi pour une semaine ISO donnée.
 */
function getDateOfISOWeek(week, year) {
  const simple = new Date(year, 0, 1 + (week - 1) * 7);
  const dow = simple.getDay();
  const ISOweekStart = simple;
  if (dow <= 4)
    ISOweekStart.setDate(simple.getDate() - simple.getDay() + 1);
  else
    ISOweekStart.setDate(simple.getDate() + 8 - simple.getDay());
  return ISOweekStart;
}