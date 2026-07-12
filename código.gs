/**
 * 1. FUNÇÃO PRINCIPAL (Adaptada para suportar anexos)
 * Esta função envia as candidaturas diárias com anexos e, no fim, limpa o acionador temporário.
 */
function enviarCandidaturasDiarias() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Linha onde começam os dados (ignora o cabeçalho)
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) return; // Folha vazia
  
  var dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6);
  var data = dataRange.getValues();
  
  // Defina o assunto exato do rascunho que criou no Gmail
  var assuntoRascunho = "Candidatura Espontânea – Marketing Digital – Pedro Chaves";
  var rascunho = obterRascunhoPorAssunto(assuntoRascunho);
  
  if (!rascunho) {
    Logger.log("Erro: Rascunho com o assunto especificado não foi encontrado.");
    return;
  }
  
  // EXTRAÇÃO DO CONTEÚDO E DOS ANEXOS DO RASCUNHO
  var mensagemModelo = rascunho.getMessage();
  var corpoModelo = mensagemModelo.getBody(); // Corrigido o nome da variável aqui
  var assuntoModelo = mensagemModelo.getSubject();
  var anexosModelo = mensagemModelo.getAttachments(); // Obtém os ficheiros anexados ao rascunho
  
  var limiteDiario = 2; // Envia no máximo 5 emails por execução (podes aumentar depois)
  var enviadosHoje = 0;
  
  for (var i = 0; i < data.length; i++) {
    var linha = data[i];
    var empresa = linha[0];
    var email = linha[1];
    var responsavel = linha[2];
    var areaFuncao = linha[3];
    var funcao = linha[4];
    var status = linha[5];
    
    // Processa apenas linhas com status 'Pendente' e com email preenchido
    if (status.toString().toLowerCase() === "pendente" && email !== "") {
      
      // Tratamento para quando não há nome do responsável
      if (!responsavel || responsavel.toString().trim() === "") {
        responsavel = "Equipa de Recursos Humanos da " + empresa;
      }
      
      // Substituição das tags dinâmicas {{ }}
      var corpoPersonalizado = corpoModelo
        .replace(/\{\{Responsável\}\}/g, responsavel)
        .replace(/\{\{Empresa\}\}/g, empresa)
        .replace(/\{\{Área\}\}/g, areaFuncao)
        .replace(/\{\{Função\}\}/g, funcao);
        
      var assuntoPersonalizado = assuntoModelo
        .replace(/\{\{Responsável\}\}/g, responsavel)
        .replace(/\{\{Empresa\}\}/g, empresa)
        .replace(/\{\{Área\}\}/g, areaFuncao)
        .replace(/\{\{Função\}\}/g, funcao);
        
      // Envio do email através do GmailApp incluindo a lista de anexos
      GmailApp.sendEmail(email, assuntoPersonalizado, "", {
        htmlBody: corpoPersonalizado,
        attachments: anexosModelo 
      });
      
      // Atualiza o status e a data na tabela
      var linhaAtual = startRow + i;
      sheet.getRange(linhaAtual, 6).setValue("Enviado");
      sheet.getRange(linhaAtual, 7).setValue(new Date());
      enviadosHoje++;
      
      if (enviadosHoje >= limiteDiario) {
        Logger.log("Limite diário de " + limiteDiario + " envios atingido.");
        break;
      }
    }
  }
  
  // CRUCIAL: Depois de enviar os e-mails, elimina o acionador para não acumular lixo
  eliminarAcionadoresTemporarios();
}

/**
 * 2. FUNÇÃO AUXILIAR ORIGINAL
 * Procura o rascunho de e-mail no teu Gmail pelo assunto.
 */
function obterRascunhoPorAssunto(assunto) {
  var drafts = GmailApp.getDrafts();
  for (var i = 0; i < drafts.length; i++) {
    if (drafts[i].getMessage().getSubject() === assunto) {
      return drafts[i];
    }
  }
  return null;
}

// =========================================================================
// SISTEMA DE AGENDAMENTO AUTOMÁTICO (PARA AS 11h01)
// =========================================================================

/**
 * 3. FUNÇÃO DE AGENDAMENTO
 * Cria o alarme interno para disparar o envio precisamente às 11h01.
 */
function agendarParaAs1101() {
  eliminarAcionadoresTemporarios();
  
  // Garante que usa o fuso horário correto da própria Folha de Cálculo
  var fusoHorario = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var hojeFormatado = Utilities.formatDate(new Date(), fusoHorario, "yyyy-MM-dd");
  
  // Cria a data alvo baseada na string do dia atual às 11:01
  var horaAlvo = new Date(hojeFormatado + "T11:01:00");
  var agora = new Date();
  
  // Se já passou das 11h01 hoje, agenda para amanhã
  if (agora > horaAlvo) {
    horaAlvo.setDate(horaAlvo.getDate() + 1);
  }
  
  ScriptApp.newTrigger("enviarCandidaturasDiarias")
           .timeBased()
           .at(horaAlvo)
           .inTimezone(fusoHorario) // Força o fuso horário correto no acionador
           .create();
           
  Logger.log("Sucesso! Próximo envio agendado para: " + horaAlvo.toString());
}

/**
 * 4. FUNÇÃO DE LIMPEZA
 * Remove os acionadores temporários criados para a função "enviarCandidaturasDiarias".
 */
function eliminarAcionadoresTemporarios() {
  var acionadores = ScriptApp.getProjectTriggers();
  for (var i = 0; i < acionadores.length; i++) {
    if (acionadores[i].getHandlerFunction() === "enviarCandidaturasDiarias") {
      ScriptApp.deleteTrigger(acionadores[i]);
    }
  }
}
