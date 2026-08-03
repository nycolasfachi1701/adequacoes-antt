/**
 * ============================================================================
 *  BACKEND DE SUGESTÕES — Portal CIOT ANTT / NSTECH
 *  Recebe os envios do formulário do portal e grava numa planilha do Google.
 *  Opcionalmente, envia um e-mail para você a cada nova sugestão.
 * ============================================================================
 *
 *  COMO CONFIGURAR (uma vez só, ~5 minutos):
 *
 *  1) Crie uma planilha nova no Google Sheets (ela será seu "painel"/banco de
 *     dados). Dê o nome que quiser.
 *
 *  2) Nessa planilha, vá em  Extensões > Apps Script.
 *     Apague todo o conteúdo que aparecer e COLE este arquivo inteiro.
 *
 *  3) Ajuste as 2 configurações abaixo:
 *        - ADMIN_TOKEN : troque por uma senha secreta só sua (qualquer texto).
 *                        Você usará ela na página admin.html.
 *        - NOTIFY_EMAIL: (opcional) seu e-mail para receber aviso de cada
 *                        nova sugestão. Deixe '' para não receber e-mails.
 *
 *  4) Clique em  Implantar > Nova implantação.
 *        - Em "Selecione o tipo" (engrenagem) escolha:  App da Web
 *        - Executar como:      Eu (seu e-mail)
 *        - Quem pode acessar:  Qualquer pessoa
 *        - Clique em Implantar e AUTORIZE o acesso quando for solicitado.
 *        - COPIE a "URL do app da Web" (termina em /exec).
 *
 *  5) Abra o index.html do portal, ache a linha:
 *        const SUGESTOES_ENDPOINT = '';
 *     e cole a URL entre as aspas. Salve. Pronto: o formulário passa a gravar
 *     na sua planilha.
 *
 *  6) Para LER os envios: abra o admin.html, cole a mesma URL e o ADMIN_TOKEN.
 *
 *  OBSERVAÇÃO: sempre que editar este script, gere uma NOVA implantação
 *  (ou "Gerenciar implantações" > editar > Nova versão) para publicar a mudança.
 * ============================================================================
 */

// ==== CONFIGURAÇÃO ====
var ADMIN_TOKEN  = 'TROQUE-ESTA-SENHA';   // senha usada no admin.html
var NOTIFY_EMAIL = '';                     // ex.: 'nycolas.fachi@nstech.com.br' (ou '' para não notificar)
var SHEET_NAME   = 'Sugestões';

// ============================================================================
//  Recebe o formulário (POST do portal)
// ============================================================================
function doPost(e){
  try{
    var dados = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    sh.appendRow([
      new Date(),
      dados.tipo     || '',
      dados.assunto  || '',
      dados.mensagem || '',
      dados.nome     || '',
      dados.empresa  || '',
      dados.email    || '',
      dados.origem   || '',
      'Novo'
    ]);

    if(NOTIFY_EMAIL){
      try{
        var corpo =
          'Tipo: '     + (dados.tipo     || '-') + '\n' +
          'Assunto: '  + (dados.assunto  || '-') + '\n\n' +
          (dados.mensagem || '') + '\n\n' +
          '--\n' +
          'Nome: '     + (dados.nome    || '-') + '\n' +
          'Empresa: '  + (dados.empresa || '-') + '\n' +
          'E-mail: '   + (dados.email   || '-') + '\n' +
          'Origem: '   + (dados.origem  || '-');
        MailApp.sendEmail(NOTIFY_EMAIL, '[Portal CIOT] Nova sugestão: ' + (dados.assunto || 'sem assunto'), corpo);
      }catch(mailErr){ /* não bloqueia a gravação se o e-mail falhar */ }
    }

    return json_({ ok: true });
  }catch(err){
    return json_({ ok: false, error: String(err) });
  }
}

// ============================================================================
//  Retorna as sugestões (GET usado pelo admin.html via JSONP, protegido por token)
// ============================================================================
function doGet(e){
  var cb = e.parameter.callback;
  if(e.parameter.token !== ADMIN_TOKEN){
    return reply_({ ok: false, error: 'unauthorized' }, cb);
  }

  var sh = getSheet_();

  // Ação de exclusão: ?action=delete&row=N (N = número da linha na planilha)
  if(e.parameter.action === 'delete'){
    var row = parseInt(e.parameter.row, 10);
    if(row >= 2 && row <= sh.getLastRow()){
      sh.deleteRow(row);
      return reply_({ ok: true, deleted: row }, cb);
    }
    return reply_({ ok: false, error: 'linha inválida' }, cb);
  }

  // Listagem
  var values = sh.getDataRange().getValues();
  var items = [];
  for(var i = 1; i < values.length; i++){
    var r = values[i];
    items.push({
      row:      i + 1, // número real da linha na planilha (cabeçalho = linha 1)
      data:     r[0],
      tipo:     r[1],
      assunto:  r[2],
      mensagem: r[3],
      nome:     r[4],
      empresa:  r[5],
      email:    r[6],
      origem:   r[7],
      status:   r[8]
    });
  }
  items.reverse(); // mais recentes primeiro
  return reply_({ ok: true, items: items }, cb);
}

// ============================================================================
//  Auxiliares
// ============================================================================
function getSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Data','Tipo','Assunto','Mensagem','Nome','Empresa','E-mail','Origem','Status']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function reply_(obj, cb){
  var s = JSON.stringify(obj);
  if(cb){
    return ContentService
      .createTextOutput(cb + '(' + s + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(obj);
}
