/*function buscaEnd() {
  var guiaAtiva = SpreadsheetApp.getActive().getSheetName();

  if (guiaAtiva == "Menu") {
    var guiaMenu = SpreadsheetApp.getActive().getSheetByName("Menu");
    var celula = guiaMenu.getActiveCell().getA1Notation();

    if (celula == "C4") {
      pesquisaCNPJs(); 
    }
  }
}*/

function pesquisaCNPJs() {

  var planilha = SpreadsheetApp.openById("INSIRA A URL DA PLANILHA AQUI"); // Planilha vinculada
  //var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaMenu = planilha.getSheetByName("INSIRA O NOME DA PAGINA DA PLANILHA AQUI");
  var scriptProperties = PropertiesService.getScriptProperties();

  var cnpjs = guiaMenu.getRange("INSIRA A CELULA DA COLUNA ONDE INICIA A CONSULTA E ONDE FINALIZA EX: 'C4:C504'").getValues(); 
  var consultasPorMinuto = 3; 
  var tempoEspera = 60000 / consultasPorMinuto;
  var tempoLimite = 6 * 60 * 1000; // 6 minutos em milissegundos
  var inicioExecucao = new Date().getTime();
  var proximoIndex = parseInt(scriptProperties.getProperty("proximoIndex")) || 0;

  var cnpj; // Declara a variável cnpj fora do loop

  for (var index = proximoIndex; index < cnpjs.length; index++) {
    cnpj = cnpjs[index][0]; // Atribui o valor do CNPJ à variável cnpj

    if (cnpj === null || cnpj === undefined) {
      guiaMenu.getRange(index + 4, 4).setValue("CNPJ não encontrado");
      continue;
    }

    var cnpjLimpo = String(cnpj).replace(/\D/g, '');

    var url = "https://receitaws.com.br/v1/cnpj/" + cnpjLimpo;
    var retorno = UrlFetchApp.fetch(url, {muteHttpExceptions: true});

    if (retorno.getResponseCode() == 200) {
      var dados = JSON.parse(retorno.getContentText());
      if (dados.status == "OK") {
        guiaMenu.getRange(index + 4, 4).setValue(dados.nome);
        guiaMenu.getRange(index + 4, 5).setValue(dados.fantasia || "Não informado");
        guiaMenu.getRange(index + 4, 6).setValue(dados.inscricao_estadual || "Não informado");
        guiaMenu.getRange(index + 4, 7).setValue(dados.opcao_pelo_simples ? "SIM" : "NÃO");
        guiaMenu.getRange(index + 4, 8).setValue(dados.situacao || "Não informado");
        guiaMenu.getRange(index + 4, 9).setValue(dados.logradouro);
        guiaMenu.getRange(index + 4, 10).setValue(dados.numero);
        guiaMenu.getRange(index + 4, 11).setValue(dados.municipio);
        guiaMenu.getRange(index + 4, 12).setValue(dados.bairro);
        guiaMenu.getRange(index + 4, 13).setValue(dados.uf);
        guiaMenu.getRange(index + 4, 14).setValue(dados.cep);
      } else {
        guiaMenu.getRange(index + 4, 4).setValue("CNPJ não encontrado");
      }
    } else {
      guiaMenu.getRange(index + 4, 4).setValue("Erro na consulta");
    }

    SpreadsheetApp.flush(); 

    if (index < cnpjs.length - 1) {
      Utilities.sleep(tempoEspera); 
    }

    // Verifica o tempo limite
    if (new Date().getTime() - inicioExecucao > tempoLimite) {
      scriptProperties.setProperty("proximoIndex", index + 1);
      ScriptApp.newTrigger("pesquisaCNPJs")
        .timeBased()
        .after(60000) 
        .create();
      return; 
    }
  }

  scriptProperties.deleteProperty("proximoIndex");
}

//Chame a função abaixo em um botão na planilha caso queira usar um botão para iniciar a consulta
function executarPesquisaCNPJs() {
  pesquisaCNPJs(); // Chama a função principal
}
