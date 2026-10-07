/**
 * Envoi des fiches BENIPOULE par mail à IMPACT PLUS.
 * Chaque fiche remplie sur le site arrive par mail avec un vrai fichier Word (.docx) et un PDF.
 */
var DESTINATAIRE = 'impactpluscabinet@gmail.com';

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var maintenant = new Date();
  var date = Utilities.formatDate(maintenant, 'Africa/Porto-Novo', 'dd/MM/yyyy à HH:mm');
  var nom = (data.filename || 'Fiche_collecte_BENIPOULE') + '_' +
            Utilities.formatDate(maintenant, 'Africa/Porto-Novo', 'yyyy-MM-dd_HH-mm');

  var pieces = [];
  try {
    pieces.push(htmlVersDocx(data.word, nom));
  } catch (err) {
    // Secours : si la conversion échoue, on joint la version Word classique
    pieces.push(Utilities.newBlob('﻿' + data.word, 'application/msword', nom + '.doc'));
  }
  if (data.pdf) {
    pieces.push(Utilities.newBlob(Utilities.base64Decode(data.pdf), 'application/pdf', nom + '.pdf'));
  }

  MailApp.sendEmail({
    to: DESTINATAIRE,
    subject: 'Nouvelle fiche BENIPOULE reçue – ' + date,
    htmlBody: '<p>Bonjour,</p><p>Une nouvelle <b>fiche de collecte BENIPOULE</b> a été remplie le ' + date +
              '.</p><p>Vous trouverez la fiche complétée en pièce jointe (Word et PDF).</p>',
    attachments: pieces
  });
  return ContentService.createTextOutput('ok');
}

/** Transforme la fiche en vrai document Word (.docx) en passant par Google Docs. */
function htmlVersDocx(html, nom) {
  var token = ScriptApp.getOAuthToken();
  var limite = 'limite_benipoule_' + new Date().getTime();
  var corps = '--' + limite + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({ name: nom, mimeType: 'application/vnd.google-apps.document' }) + '\r\n' +
    '--' + limite + '\r\n' +
    'Content-Type: text/html; charset=UTF-8\r\n\r\n' +
    html + '\r\n' +
    '--' + limite + '--';

  var creation = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + limite,
    payload: Utilities.newBlob(corps).getBytes(),
    headers: { Authorization: 'Bearer ' + token }
  });
  var id = JSON.parse(creation.getContentText()).id;

  try {
    var docx = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id +
      '/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document', {
      headers: { Authorization: 'Bearer ' + token }
    }).getBlob();
    return docx.setName(nom + '.docx');
  } finally {
    DriveApp.getFileById(id).setTrashed(true); // on ne garde pas de copie dans Drive
  }
}

/** À lancer une seule fois à la main : demande les autorisations (aucun mail n'est envoyé). */
function autoriser() {
  DriveApp.getRootFolder();
  UrlFetchApp.getRequest('https://www.googleapis.com/drive/v3/about');
  MailApp.getRemainingDailyQuota();
  Logger.log('Autorisations OK');
}
