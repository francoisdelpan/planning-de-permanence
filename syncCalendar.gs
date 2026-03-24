/**
 * Ne synchronise QUE les modifs de l'agenda de l'utilisateur.
 */
function syncGoogleCalendarsForUser() {
  const me    = Session.getEffectiveUser().getEmail();
  const local = me.split('@')[0];
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const db    = ss.getSheetByName(DB_SHEET);
  if (!db) return { success: 0, failed: 0 };

  const data = db.getDataRange().getValues();

  // 1) On filtre en amont
  const rowsToProcess = data
    .slice(1)
    .map((r, i) => ({ row: r, idx: i + 2 }))
    .filter(({ row }) =>
      String(row[COL_STATUS]).trim() === "Non synchronisé"
      && sanitizeName(row[COL_NAME]) === local
    );

  // Si rien à faire, on quitte proprement
  if (rowsToProcess.length === 0) {
    return { success: 0, failed: 0 };
  }

  // 2) Initialisations
  const rowsToDelete = [];
  let success = 0;
  let failed  = 0;

  rowsToProcess.forEach(({ row, idx }) => {
    const dateCode = String(row[COL_DATECODE]);
    const newValue = row[COL_NEWVALUE];

    // Reconstruction de la date
    const [y,m,d] = dateCode.match(/(\d{4})(\d{2})(\d{2})/).slice(1).map(Number);
    const dayStart = new Date(y, m-1, d);
    const pStart   = new Date(y, m-1, d, PERMANENCE_START_HOUR);
    const pEnd     = new Date(y, m-1, d, PERMANENCE_END_HOUR);

    // Parsing de la colonne C (mix {calId,eventId} & strings)
    let arr = [];
    try { arr = JSON.parse(row[COL_CALENDAR]); } catch(e){}

    // Séparation anciens liens et liste de calendriers
    const oldLinks = [];
    const calIds   = [];
    arr.forEach(item => {
      if (item && item.calId && item.eventId) {
        oldLinks.push(item);
        calIds.push(item.calId);
      } else if (typeof item === 'string') {
        calIds.push(item);
      }
    });

    // 3) Suppression d'abord
    oldLinks.forEach(({ calId, eventId }) => {
      try {
        const ev = CalendarApp.getCalendarById(calId).getEventById(eventId);
        if (ev) ev.deleteEvent();
      } catch(err) {
        Logger.log(`❌ Suppression ${calId}/${eventId} : ${err}`);
        failed++;
      }
    });

    // 4) Si on a retiré le créneau, on planifie la suppression de la ligne
    if (!newValue) {
      rowsToDelete.push(idx);
      success++;
      return;   // on passe à la ligne suivante
    }

    // 5) Sinon création des nouveaux events
    const newLinks = [];
    calIds.forEach(calId => {
      const cal = CalendarApp.getCalendarById(calId);
      if (!cal) {
        Logger.log(`⚠️ Impossible d’accéder au calendrier ${calId} pour ${row[COL_NAME]}`);
        failed++;
        return;
      }
      try {
        const cal = CalendarApp.getCalendarById(calId);
        let ev = newValue === "P"
          ? cal.createEvent(`Permanence ${row[COL_NAME]}`, pStart, pEnd)
          : cal.createAllDayEvent(`${newValue} ${row[COL_NAME]}`, dayStart);
        newLinks.push({ calId, eventId: ev.getId() });
      } catch(err) {
        SpreadsheetApp.getUi().alert(`❌ Création ${calId} : ${err}`);
        failed++;
      }
    });

    /*// 6) Mise à jour de la ligne
    db.getRange(idx, COL_CALENDAR  +1).setValue(JSON.stringify(newLinks));
    db.getRange(idx, COL_OLDVALUE  +1).setValue(newValue);
    db.getRange(idx, COL_NEWVALUE  +1).setValue("");
    db.getRange(idx, COL_STATUS    +1).setValue("Synchronisé");
    success++;*/
    // 6) Mise à jour de la ligne (mettre à jour uniquement si on a au moins 1 lien)
    if (newLinks.length > 0) {
      db.getRange(idx, COL_CALENDAR+1).setValue(JSON.stringify(newLinks));
      db.getRange(idx, COL_OLDVALUE+1).setValue(newValue);
      db.getRange(idx, COL_NEWVALUE+1).setValue("");
      db.getRange(idx, COL_STATUS+1).setValue("Synchronisé");
      success++;
    } else {
      // on peut aussi marquer "Échec" ou laisser "Non synchronisé"
      db.getRange(idx, COL_STATUS+1).setValue("Non synchronisé");
      failed++;
      Logger.log(`Aucun événement créé pour ${row[COL_NAME]}, on conserve les calIds`);
    }

  });

  // 7) Suppression des lignes marquées, de bas en haut
  rowsToDelete
    .sort((a,b) => b - a)
    .forEach(r => db.deleteRow(r));

  // 8) On renvoie le bilan pour le client
  return { success, failed };
}

function sanitizeName(name) {
  return name
    .normalize('NFD')                       // décompose les accents
    .replace(/[\u0300-\u036f]/g, '')        // supprime les diacritiques
    .toLowerCase()                          // en minuscules
    .replace(/[^a-z0-9]+/g, '_')            // tout non-alphanum → '_'
    .replace(/^_|_$/g, '');                 // retire '_' en début/fin
}