function PageBatchInputCessao() {
  
  var pagCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input Cessão');
  
  /*Retorna a posição da última linha que possui conteúdo e posiciona na 1ª linha para inserção do próximo dado. */  
  var lastRow_Cessao = pagCessao.getLastRow();
  var lastColumn_Cessao = pagCessao.getLastColumn();
  /*
  *  Retorna o intervalo com a célula superior esquerda nas coordenadas fornecidas com o número de linhas e colunas.
  */  
  var dataRange = pagCessao.getRange(5, 2, lastRow_Cessao-4, lastColumn_Cessao);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  return data;
  
}
