function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Permanence')
    .addItem("🖨️ Imprimer la feuille de perm.", "printExternalSheet2")
    .addSeparator()
    .addItem('🔄  Synchroniser les Agenda', 'openSyncModal')
    .addItem('🌐 Modifier les Agendas Google', 'openCalendarModal')
    .addItem('👀 Voir les créneaux non synchronisés par Manager', 'showUnsyncedCounts')
    .addSeparator()
    .addItem('📆 Créer un  nouveau planning', 'openPlanningGenerator')
    .addItem('🧹 Supprimer un planning obsolète', 'openDeleteSheetModal')
    .addSeparator()
    //.addItem("🔐 Activer les autorisations", "authorizeCalendar")
    .addItem('📘 Aide utilisateur', 'openHelpPDF')
    .addToUi();
}

function openPlanningGenerator() {
  const maxWeeks = MAX_WEEKS;

  const html = HtmlService
    .createTemplateFromFile('modal_planning');
  html.maxWeeks = maxWeeks;

  const output = html.evaluate().setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(output, 'Générer un planning');
}

function openCalendarModal() {
  const html = HtmlService.createHtmlOutputFromFile("modal_calendar")
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Gestion des Calendriers");
}

function openDeleteSheetModal() {
  const html = HtmlService.createHtmlOutputFromFile("modal_delete")
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Supprimer un planning");
}

function openSyncModal() {
  const html = HtmlService.createHtmlOutputFromFile("modal_sync")
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Synchroniser les Agendas");
}

function authorizeCalendar() {
  try {
    // Appel trivial : liste les 1ers events du calendrier principal
    const opt = {maxResults: 1};
    const events = Calendar.Events.list('primary', opt);
    SpreadsheetApp.getUi().alert(
      '✅ Autorisation réussie ! Vous pouvez maintenant synchroniser les disponibilités.'
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      '❌ Échec lors de l’autorisation : ' + err.message
    );
  }
}

function showUnsyncedCounts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEET);
  if (!sheet) return ui.alert('Feuille DB introuvable: ' + DB_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ui.alert('Aucune donnée DB');

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const counts = getUnsyncedCounts(rows);
  let msg = 'Créneaux non synchronisés par utilisateur:\n';
  for (const [email, info] of Object.entries(counts)) {
    msg += `${info.name} (${email}): ${info.count}\n`;
  }
  ui.alert(msg);
}

function printExternalSheet1() {
  const fileId = "1PcL0lBvQyC4vLITVLLjtXVxCOwKECRLP"; // ID du fichier externe
  const gid = "363877205"; // GID de la feuille PRINT (à ajuster)

  const printUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf` +
    `&gid=${gid}` +
    `&portrait=false` +      // Paysage
    `&fitw=true` +           // Ajuster à la largeur
    `&scale=4` +             // Échelle = Ajuster à la page
    `&sheetnames=false` +
    `&printtitle=false` +
    `&pagenumbers=true` +
    `&gridlines=false` +
    `&fzr=false`;            // Ne pas répéter les lignes figées

  // Ouvre un nouvel onglet vers le lien de l'aperçu
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${printUrl}", "_blank");google.script.host.close();</script>`
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, "Téléchargement de l'impression");
}

function printExternalSheet2() {
  const fileId = "1PcL0lBvQyC4vLITVLLjtXVxCOwKECRLP"; // ID du fichier externe
  const gid = "1028013614"; // GID de la feuille PRINT (à ajuster)

  const printUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf` +
  `&gid=${gid}` +
  `&portrait=true` +         // Portrait
  `&fitw=true` +             // Ajuster à la largeur
  `&scale=2` +               // Échelle adaptée (~94%)
  `&size=A4` +               // Papier standard
  `&sheetnames=false` +
  `&printtitle=false` +
  `&pagenumbers=true` +
  `&gridlines=false` +
  `&fzr=false` +
  `&top_margin=0.25` +
  `&bottom_margin=0.25` +
  `&left_margin=0.25` +
  `&right_margin=0.25`;

  // Ouvre un nouvel onglet vers le lien de l'aperçu
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${printUrl}", "_blank");google.script.host.close();</script>`
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, "Téléchargement de l'impression");
}

/*function openHelpPDF() {
  const pdfUrl = "https://docs.google.com/document/d/1r0rJJrcQY34OhagZYdyxYaC2RTh3jfNnDQJKUAP61Eg/edit?tab=t.0#heading=h.e6rhlzhlha99"; // 👈 mets ici ton vrai lien Drive partagé

  const html = HtmlService.createHtmlOutput(
    `<script>
      window.open("${pdfUrl}", "_blank");
      google.script.host.close();
    </script>`
  ).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, "Redirection...");
}*/

function openHelpPDF() {
  const pdfUrl = "https://drive.google.com/file/d/1_HMvGkGiY2Lz8sqyCwTODw3xN5NbqIs2/view?usp=sharing"; // 👈 mets ici ton vrai lien Drive partagé

  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open("${pdfUrl}", "_blank");
      google.script.host.close();
    </script>
  `).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, "Ouverture de l’aide");
}