/**
 * Code à coller dans Google Apps Script (Extensions > Apps Script de ton Google Sheet).
 * Il reçoit les réponses du quiz et ajoute une ligne dans la feuille.
 */
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // Première fois : on écrit les titres de colonnes
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Score (%)', 'Réponses']);
  }

  var data = JSON.parse(e.postData.contents);

  // On met chaque réponse sous la forme "Question : Réponse"
  var reponses = (data.answers || [])
    .map(function (x) { return x.q + ' → ' + x.a; })
    .join('\n');

  sheet.appendRow([new Date(), data.score, reponses]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
