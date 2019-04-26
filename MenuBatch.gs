// Use this code for Google Docs, Slides, Forms, or Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Mostara Menu para DocumentApp or SlidesApp or FormApp.
      .createMenu('Menu Id Baixas')
      .addItem('1 - Passo a Passo - Ident. de Baixa:', 'IdBaixa') 
      .addItem('2 - Atualizar Baixas :', 'AtualizaBaixas') 
      .addItem('3 - Passo a Passo - Identificar o Cógigo Cliente - Processo Manual:', 'CodCliente') 
      .addItem('4 - Gerar Relatório de Recebimento', 'RelatorioRecebimento')
      .addItem('5 - Gerar Batch Input', 'BatchInput')      
      .addItem('6 - Gerar TXT da Batch Input', 'GeraTxt')
      .addItem('7 - Gerar Batch Input Cessão', 'BatchInputCessao')      
      .addItem('8 - Gerar TXT Cessãp da Batch Input Cessão', 'GeraTxtCessao')
  
      .addToUi();
  dialog.setPopupPosition(100, 100).setSize(500, 500);
      
}

function IdBaixa() 
{
    var html = HtmlService.createHtmlOutputFromFile('PassoaPasso');
     html.setWidth(900).setHeight(850);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, 'IMPORTANTE... Leia o Passo a Passo.');
  
        
}
function CodCliente() 
{
    var html = HtmlService.createHtmlOutputFromFile('IdCodCliente');
     html.setWidth(900).setHeight(850);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showModalDialog(html, 'IMPORTANTE... Leia o Passo a Passo.');
  
        
}