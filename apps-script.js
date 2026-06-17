// Google Apps Script — Cola isso em script.google.com
// depois de criar a planilha no Google Sheets

const EMAIL_NOTIFICACAO = 'recria.agenciamkt@gmail.com';
const NOME_PLANILHA = 'Leads Recria';

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName(NOME_PLANILHA);

    if (!aba) {
      aba = ss.insertSheet(NOME_PLANILHA);
      aba.appendRow(['Data', 'Nome', 'Negócio', 'WhatsApp', 'E-mail', 'Momento', 'Mensagem']);
      aba.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    const agora = new Date();
    const dataFormatada = Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

    aba.appendRow([
      dataFormatada,
      dados.nome || '',
      dados.negocio || '',
      dados.whatsapp || '',
      dados.email || '',
      dados.momento || '',
      dados.mensagem || ''
    ]);

    const assunto = `🔔 Novo lead Recria — ${dados.nome} (${dados.momento})`;
    const corpo = `
Novo lead entrou pela LP da Recria!

Nome: ${dados.nome}
Negócio: ${dados.negocio}
WhatsApp: ${dados.whatsapp}
E-mail: ${dados.email}
Momento: ${dados.momento}
Mensagem: ${dados.mensagem || '(não preenchido)'}

Data: ${dataFormatada}

Acesse a planilha completa:
${ss.getUrl()}
    `.trim();

    MailApp.sendEmail(EMAIL_NOTIFICACAO, assunto, corpo);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Recria Lead API — OK');
}
