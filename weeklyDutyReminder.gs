/**
 * Envoie chaque dimanche soir à 20h00 un email aux utilisateurs
 * listant leurs permanences de la semaine suivante.
 */

/**
 * Normalise un nom pour générer la partie locale de l'email.
 */
function sanitizeName(name) {
  return name
    .normalize('NFD')                       // décompose les accents
    .replace(/[̀-\u036f]/g, '')        // supprime les diacritiques
    .toLowerCase()                          // en minuscules
    .replace(/[^a-z0-9]+/g, '_')            // tout non-alphanum → '_'
    .replace(/^_|_$/g, '');                 // retire '_' en début/fin
}

/**
 * Retourne "Jour jj/MM" à partir de "jj/MM" et de l'année de référence.
 * On prend l'année du lundi de la semaine (safe même en fin d'année).
 */
function formatDayLabel(dateStr, refYear) {
  const [dd, mm] = dateStr.split('/').map(x => parseInt(x, 10));
  const d = new Date(refYear, mm - 1, dd);
  return `${DAYS_FR[d.getDay()]} ${dateStr}`;
}

/**
 * Construit l'adresse email à partir du nom.
 * Lance une erreur si la partie locale est invalide.
 */
function sanitizeEmail(name) {
  const localPart = sanitizeName(name);
  if (!localPart) {
    throw new Error('sanitizeEmail: nom vide ou invalide -> ' + name);
  }
  return `${localPart}@franchise.carrefour.com`;
}

/**
 * Fonction principale pour envoyer les emails hebdomadaires.
 */
function sendWeeklyDutyReminder() {
  // Pour debug, définir 'YYYY-MM-DD', sinon null
  const debugDate = null;
  const today = debugDate ? new Date(debugDate) : new Date();

  // 1. Calcul du prochain dimanche
  const offset    = (7 - today.getDay()) % 7;
  const nextSunday = new Date(today);
  nextSunday.setDate(today.getDate() + offset);
  Logger.log("nextSunday : " + nextSunday);

  // 2. Détermination de la plage semaine (Lundi -> Dimanche)
  const msPerDay = 24 * 60 * 60 * 1000;
  const monday   = new Date(nextSunday.getTime() + msPerDay);
  const sunday   = new Date(monday.getTime() + 6 * msPerDay);
  Logger.log("monday : " + monday + " sunday : " + sunday);

  // 3. Calcul ISO (semaine et année)
  function getISOWeekAndYear(d) {
    const dateUTC = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum  = dateUTC.getUTCDay() || 7;
    dateUTC.setUTCDate(dateUTC.getUTCDate() + 4 - dayNum);
    const year = dateUTC.getUTCFullYear();
    const week = Math.ceil(( (dateUTC - new Date(Date.UTC(year, 0, 1))) / msPerDay + 1 ) / 7);
    return { week, year };
  }
  const iso        = getISOWeekAndYear(monday);
  const targetComp = iso.year * 100 + iso.week;
  Logger.log("targetComp : " + targetComp);

  // 4. Recherche de la feuille couvrant la semaine cible
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  let sheetFound = null;
  for (const s of ss.getSheets()) {
    const m = s.getName().match(/^S(\d{1,2})\s+(\d{4})\s*-\s*S(\d{1,2})\s+(\d{4})$/);
    if (m) {
      const startComp = parseInt(m[2], 10) * 100 + parseInt(m[1], 10);
      const endComp   = parseInt(m[4], 10) * 100 + parseInt(m[3], 10);
      if (targetComp >= startComp && targetComp <= endComp) {
        sheetFound = s;
        break;
      }
    }
  }
  if (!sheetFound) {
    MailApp.sendEmail(
      ADMIN_EMAIL,
      '[ERREUR] Planning non trouvé',
      `Aucune feuille trouvée pour la semaine ${iso.week} ${iso.year}`
    );
    return;
  }
  const sheet = sheetFound;

  // 5. Lecture de la ligne des dates
  const lastCol = sheet.getLastColumn();
  const datesRow = sheet.getRange(LINE_DATE, 1, 1, lastCol).getValues()[0];
  const dateMap = {};
  datesRow.forEach((v, i) => {
    if (v) dateMap[v.toString().slice(0,5)] = i + 1;
  });

  // 6. Construction du tableau des dates ('jj/MM')
  /*const pad = n => ('0' + n).slice(-2);
  const slotDates = [];
  for (let d = new Date(monday); d <= sunday; d.setDate(d.getDate() + 1)) {
    slotDates.push(`${pad(d.getDate())}/${pad(d.getMonth()+1)}`);
  }*/
  const pad = n => ('0' + n).slice(-2);
  const slotDates = [];
  for (let d = new Date(monday); d <= sunday; d.setDate(d.getDate() + 1)) {
    const dateObj = new Date(d);  // copie
    const dateStr = `${pad(dateObj.getDate())}/${pad(dateObj.getMonth()+1)}`;
    slotDates.push({ dateStr, dateObj });
  }
  Logger.log("slotDates : " + slotDates);

  // 7. Collecte des permanences par utilisateur
  const permsByUser = {};
  /*for (const dateStr of slotDates) {
    const col = dateMap[dateStr];
    if (!col) {
      MailApp.sendEmail(
        ADMIN_EMAIL,
        '[ERREUR] Date manquante dans le planning',
        `La date ${dateStr} est introuvable sur ${sheet.getName()}`
      );
      continue;
    }
    const numRows = sheet.getLastRow();
    const values  = sheet.getRange(LINE_TRAME_USER_START, col, numRows - LINE_TRAME_USER_START + 1).getValues();
    values.forEach((rowArr, idx) => {
      if (rowArr[0] === 'P') {
        const row  = LINE_TRAME_USER_START + idx;
        const name = sheet.getRange(row, COL_TRAME_USER).getValue();
        let email;
        try {
          email = sanitizeEmail(name);
          Logger.log("email : " + email);
        } catch (e) {
          MailApp.sendEmail(
            ADMIN_EMAIL,
            '[ERREUR] Email invalide',
            `Impossible de parser l'email de ${name}`
          );
          return;
        }
        if (!permsByUser[email]) permsByUser[email] = { name, dates: [] };
        permsByUser[email].dates.push(dateStr);
      }
    });
    Logger.log(`permsByUser after ${dateStr}: ` + JSON.stringify(permsByUser));
  }*/
  for (const { dateStr, dateObj } of slotDates) {
    const col = dateMap[dateStr];
    if (!col) {
      MailApp.sendEmail(ADMIN_EMAIL,
        '[ERREUR] Date manquante',
        `La date ${dateStr} est introuvable sur ${sheet.getName()}`
      );
      continue;
    }
    const numRows = sheet.getLastRow();
    const values = sheet.getRange(LINE_TRAME_USER_START, col, numRows - LINE_TRAME_USER_START + 1).getValues();
    values.forEach((rowArr, idx) => {
      if (rowArr[0] === 'P') {
        const row  = LINE_TRAME_USER_START + idx;
        const name = sheet.getRange(row, COL_TRAME_USER).getValue();
        let email;
        try {
          email = sanitizeEmail(name);
        } catch (e) {
          MailApp.sendEmail(ADMIN_EMAIL,
            '[ERREUR] Email invalide',
            `Impossible de parser l'email pour '${name}'`
          );
          return;
        }
        if (!permsByUser[email]) permsByUser[email] = { name, dates: [] };
        // stocke l'objet complet pour avoir dateObj plus tard
        permsByUser[email].dates.push({ dateStr, dateObj });
      }
    });
    Logger.log(`permsByUser after ${dateStr}: ` + JSON.stringify(permsByUser));
  }

  // 8. Envoi des emails
  for (const [email, info] of Object.entries(permsByUser)) {
    const subject = `[INFO] Vos permanences de la semaine ${iso.week}`;

    const htmlBody = `
      <p>Bonjour ${info.name},</p>
      <p>Vous êtes de permanence la semaine prochaine :</p>
      <ul>
        ${info.dates.map(({ dateStr, dateObj }) =>
          `<li>${DAYS_FR[dateObj.getDay()]} ${dateStr}</li>`
        ).join('')}
      </ul>
      <p>Lien du <a href="${PLANNING_URL}" target="_blank">planning de permanence</a>.</p>
      <p>Bien à vous,<br>Bonne semaine.</p>
    `;

    MailApp.sendEmail({
      to:       email,
      subject:  `[INFO] Vos permanences de la semaine ${iso.week}`,
      htmlBody: htmlBody,
      name:     (email == 'matthieu_hamel@franchise.carrefour.com') ? 'Napoléon Bonarpart, 1er du Nom' : 'SimK Retail Perm Admin',
      replyTo:  ADMIN_EMAIL
    });
  }
}

/**
 * Configurez un trigger dans l'éditeur :
 *   Événement -> Time-driven -> Every week -> Sunday -> 20:00
 */