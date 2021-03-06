function PageBatchInput() {
  
  var pagBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input');
  
  /*Retorna a posição da última linha que possui conteúdo e posiciona na 1ª linha para inserção do próximo dado. */  
  var lastRow_Layout = pagBase.getLastRow();
  var lastColumn_Layout = pagBase.getLastColumn();
  /*
  *  Retorna o intervalo com a célula superior esquerda nas coordenadas fornecidas com o número de linhas e colunas.
  */  
  var dataRange = pagBase.getRange(5, 2, lastRow_Layout-4, lastColumn_Layout);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  return data;
  
}