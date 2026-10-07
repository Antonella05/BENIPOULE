/**
 * Script d'envoi des fiches BENIPOULE par mail à IMPACT PLUS.
 * À coller dans https://script.google.com (connecté avec impactpluscabinet@gmail.com),
 * puis à déployer en « Application Web » (Exécuter en tant que : moi / Accès : Tout le monde).
 */
var DESTINATAIRE = 'impactpluscabinet@gmail.com';

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var date = Utilities.formatDate(new Date(), 'Africa/Porto-Novo', 'dd/MM/yyyy à HH:mm');
  var nom = (data.filename || 'Fiche_collecte_BENIPOULE') + '_' +
            Utilities.formatDate(new Date(), 'Africa/Porto-Novo', 'yyyy-MM-dd_HH-mm');

  var pieces = [Utilities.newBlob(data.word, 'application/msword', nom + '.doc')];
  if (data.pdf) {
    pieces.push(Utilities.newBlob(Utilities.base64Decode(data.pdf), 'application/pdf', nom + '.pdf'));
  }

  MailApp.sendEmail({
    to: DESTINATAIRE,
    subject: 'Nouvelle fiche BENIPOULE reçue – ' + date,
    htmlBody: '<p>Bonjour,</p><p>Une nouvelle <b>fiche de collecte BENIPOULE</b> a été remplie le ' + date +
              '.</p><p>Vous trouverez la fiche en pièce jointe (Word et PDF).</p>',
    attachments: pieces
  });
  return ContentService.createTextOutput('ok');
}

/** À lancer une fois à la main pour autoriser l'envoi de mails et tester. */
function testEnvoi() {
  doPost({ postData: { contents: JSON.stringify({
    filename: 'TEST_Fiche_BENIPOULE',
    word: '<html><body><h1>Test</h1><p>Si vous recevez ce mail, tout fonctionne.</p></body></html>',
    pdf: ''
  }) } });
}
