/*
  Direitos Autorais (c) 2025 Renata Veras Venturim

  Licença de Uso com Atribuição Obrigatória

  Este script foi desenvolvido por Renata Veras Venturim para automatizar o envio de e mails, com objetivo de melhorar o acesso e a transparência das informações no setor.

  Fica autorizada a utilização, cópia e redistribuição deste código, com ou sem modificações,
  desde que mantidos os créditos da autora de forma visível e permanente, tanto no código 
  quanto em qualquer sistema, dashboard ou relatório que utilize este script direta ou 
  indiretamente.

  É expressamente proibida a remoção ou ocultação do nome da autora, bem como a utilização 
  do código sem atribuição clara de autoria.

  Este software é fornecido "no estado em que se encontra", sem garantias de qualquer tipo, 
  expressas ou implícitas. O uso é de responsabilidade do usuário.

  Autoria e Desenvolvimento: Renata Veras Venturim
  Ano: 2025
*/


function enviarEmailsSaldo() {
 
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Deseja enviar e-mail para todos agora?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var nomeDaAba = "Página1";

    // Obtenha a planilha ativa pelo nome
    //var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeDaAba);
    var planilha = SpreadsheetApp.openById('1mtWTEzH6ZU27OUAZvjdjHDpACt7MOgUo4ALa5FSU6Yw').getSheetByName(nomeDaAba);

    // Verifique se a planilha foi encontrada
    if (planilha) {
      // Obtenha os dados na coluna E
      var dadosColunaE = planilha.getRange("E:E").getValues();

    // Inicialize uma array para armazenar todas as linhas que atendem ao critério
      var linhasEncontradas = [];

      // Percorra todas as linhas
      for (var i = 0; i < dadosColunaE.length; i++) {
        if (dadosColunaE[i][0].toLowerCase()) {
          // Adicione a linha à array de linhas encontradas
          linhasEncontradas.push(i + 1);
        }
      } 

      // Percorra todas as linhas encontradas da array linhasEncontradas
      for (var j = 0; j < linhasEncontradas.length; j++) {
        var linhaEncontrada = linhasEncontradas[j];

        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Página1');
        var startRow = linhaEncontrada+1; // início da linha de dados
        var numRows = 1; // número total de linhas de dados (apenas a primeira linha)
        var dataRange = sheet.getRange(startRow, 4, numRows, 14); // intervalo de dados (colunas E até N, começando da linha 1)
        var data = dataRange.getValues(); // valores das células do intervalo de dados
        var subject = 'EXTRATO FONTE ARRECADAÇÃO PRÓPRIA ( ANTIGA 250 )- %s';
        var body = '<html><body>Prezados(as), da %s<br><br>Encaminhamos abaixo o saldo disponível do recurso de Arrecadação Própria (Fonte 250) referente ao seu respectivo Programa de Pós-Graduação:<br><br><b>SALDO: R$%s</b><br><br>Informamos que o montante encontra-se disponível para utilização e financiamento das seguintes despesas:<br><b>   1. Material Permanente - (itens de capital), incluindo licença vitalícia de software.<br>   2. Material de Consumo - (itens de custeio), prestação de serviço de pessoa jurídica, pagamentos de taxas, auxílio financeiro a estudantes e a pesquisadores, diárias, passagens e bolsas.</b><br><br>Informamos que o saldo poderá ser acompanhado via link: https://lookerstudio.google.com/reporting/562d698e-6fce-4c22-b2bf-315df82dbf8f <br><br>Permanecemos à disposição para eventuais esclarecimentos.<br><br>Atenciosamente,<br>Coordenação de Administração Financeira<br>PROPPI - Pró-Reitoria de Pesquisa Pós-Graduação e Inovação<br>Rua Miguel de Frias, 9, 3º andar - Icaraí - Niterói - RJ - 24220-900</body></html>';

        for (var k = 0; k < data.length; k++) {
          var row = data[k];  // para obter cada linha
          var nomeCurso = row[1]; // valor da coluna E
          var valor = row[8]; // valor da coluna L   
          var emailAddress = row[9]; // valor da coluna M
          var competencia =  planilha.getRange("M2").getValue();// valor da coluna M linha 2

          Logger.log("Linha: " + startRow + ", e-mail: " + emailAddress);

          // Verifica se as células estão vazias antes de enviar o e-mail
      if (valor && emailAddress) {
        // Transforma todos os e-mails em minúsculo e separa por vírgula
        var emailList = emailAddress
          .split(",")
          .map(e => e.trim()) // remove espaços extras
          .filter(e => e && e.toLowerCase() !== "compras.proppi@id.uff.br"); // remove o que for igual a compras

        if (emailList.length > 0) {
          var message = body.replace('%s', nomeCurso).replace('%s', valor);
          subject = subject.replace('%s', competencia);
          MailApp.sendEmail({
            to: emailList.join(","),
            subject: subject,
            htmlBody: message,
          });
        } else {
          Logger.log("Nenhum destinatário válido na linha " + startRow + " (todos bloqueados).");
        }
      }


        }
      } 
      ui.alert("E-mails enviados com sucesso!");
    }
    else{
      ui.alert("Planilha ou aba não encontrada. Verifique os nomes da planilha e da aba.");
    } 
    }  
}




