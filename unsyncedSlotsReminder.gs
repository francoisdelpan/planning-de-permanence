/**
 * Extrait les créneaux non synchronisés et compte par utilisateur.
 * @param {Array<Array>} rows Données DB colonnes A-F
 * @returns {Object} mapping email -> {name, count}
 */
function getUnsyncedCounts(rows) {
  const unsynced = rows.filter(r => r[5] === 'Non synchronisé');
  Logger.log('getUnsyncedCounts: total unsynced=' + unsynced.length);
  const counts = unsynced.reduce((acc, r) => {
    const userName = r[1];
    let email;
    try {
      email = sanitizeEmail(userName);
    } catch (e) {
      Logger.log('sanitizeEmail error for ' + userName);
      MailApp.sendEmail(
        ADMIN_EMAIL,
        '[ERROR] Email invalide DB',
        `Impossible de parser l\'email pour '${userName}'`
      );
      return acc;
    }
    if (!acc[email]) acc[email] = { name: userName, count: 0 };
    acc[email].count++;
    return acc;
  }, {});
  Logger.log('getUnsyncedCounts result=' + JSON.stringify(counts));
  return counts;
}

/**
 * Envoie chaque lundi à 08h00 un rappel des créneaux non synchronisés.
 */
function sendUnsyncedSlotsReminder() {
  const debugDate = null;
  const today     = debugDate ? new Date(debugDate) : new Date();
  Logger.log('sendUnsyncedSlotsReminder start – '+today);

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(DB_SHEET);
  if (!sheet) { Logger.log('DB sheet missing'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No DB data'); return; }

  const rows   = sheet.getRange(2,1,lastRow-1,6).getValues();
  Logger.log('Fetched '+rows.length+' rows');
  const counts = getUnsyncedCounts(rows);

  for (const [email,info] of Object.entries(counts)) {
    Logger.log('Reminder to '+email+' count='+info.count);
    const html = `<p>Bonjour ${info.name},</p>
      <p>Vous avez ${info.count} créneau(s) non synchronisé(s).</p>
      <p><a href="${PLANNING_URL}" target="_blank">Cliquez ici</a> puis utilisez le menu <b>📅 Permanence > 🔄 Synchroniser les Agenda.</b></p>`;
    MailApp.sendEmail({to:email,subject:'[INFO] Sync en attente',htmlBody:html,
      name:'SimK Retail Perm Admin',replyTo:ADMIN_EMAIL});
  }
  Logger.log('sendUnsyncedSlotsReminder end');
}

/**
 * Déclencheurs à configurer :
 *   - sendWeeklyPermEmails()     : Hebdomadaire, Dimanche à 20h00
 *   - sendUnsyncedSlotsReminder(): Hebdomadaire, Lundi à 08h00
 */