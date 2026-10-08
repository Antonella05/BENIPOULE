var DESTINATAIRE = 'impactpluscabinet@gmail.com';
var TYPE_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var now = new Date();
  var date = Utilities.formatDate(now, 'Africa/Porto-Novo', 'dd/MM/yyyy à HH:mm');
  var nom = 'Fiche_collecte_BENIPOULE_' + Utilities.formatDate(now, 'Africa/Porto-Novo', 'yyyy-MM-dd_HH-mm');
  var pieces = [];
  try {
    pieces.push(versDocx(data.word, nom));
  } catch (err) {
    pieces.push(Utilities.newBlob(data.word, 'application/msword', nom + '.doc'));
  }
  if (data.pdf) {
    pieces.push(Utilities.newBlob(Utilities.base64Decode(data.pdf), 'application/pdf', nom + '.pdf'));
  }
  var noms = [];
  (data.files || []).forEach(function (f) {
    pieces.push(Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', f.name));
    noms.push(f.name);
  });
  var liste = noms.length
    ? '<br><br>Documents joints par le répondant (' + noms.length + ') :<br>• ' + noms.join('<br>• ')
    : '<br><br>Aucun document joint par le répondant.';
  MailApp.sendEmail({
    to: DESTINATAIRE,
    subject: 'Nouvelle fiche BENIPOULE reçue – ' + date,
    htmlBody: 'Bonjour,<br><br>Une nouvelle fiche de collecte BENIPOULE a été remplie le ' + date + '.<br>La fiche est en pièce jointe (Word et PDF).' + liste,
    attachments: pieces
  });
  return ContentService.createTextOutput('ok');
}

function versDocx(html, nom) {
  var auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var b = 'benipoule' + Date.now();
  var corps = '--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'
    + JSON.stringify({ name: nom, mimeType: 'application/vnd.google-apps.document' })
    + '\r\n--' + b + '\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n'
    + html + '\r\n--' + b + '--';
  var rep = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + b,
    payload: Utilities.newBlob(corps).getBytes(),
    headers: auth
  });
  var id = JSON.parse(rep.getContentText()).id;
  var url = 'https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=' + encodeURIComponent(TYPE_DOCX);
  var docx = UrlFetchApp.fetch(url, { headers: auth }).getBlob().setName(nom + '.docx');
  DriveApp.getFileById(id).setTrashed(true);
  return docx;
}

function autoriser() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.getRequest('https://www.googleapis.com');
  Logger.log('Autorisations OK');
}
