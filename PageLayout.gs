function PageLayout() {
  
   var pagLayout = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout');
  
  /*Retorna a posição da última linha que possui conteúdo e posiciona na 1ª linha para inserção do próximo dado. */  
  var lastRow_Layout = pagLayout.getLastRow();
  var lastColumn_Layout = pagLayout.getLastColumn();
  /*
  *  Retorna o intervalo com a célula superior esquerda nas coordenadas fornecidas com o número de linhas e colunas.
  */  
  var dataRange = pagLayout.getRange(5, 1, lastRow_Layout-4, lastColumn_Layout);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  return data;
  
}