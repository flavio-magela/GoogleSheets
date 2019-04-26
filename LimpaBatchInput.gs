function LimpaBatchInput() {
  
  var PagBatch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input');
  var inicioLinhabatchInput = 5;  
  var inicioColunabatchInput = 2;
  
  
  // instancia as planilhas
  var batchInput = new PageBatchInput();
  
  try{
    
     //funciona também
  
      for(i in batchInput){
       
        var inicioColunabatchInput = 2;
        var colunabatchInput = batchInput[i];
        var lenthbatchInput = colunabatchInput.length;
        var k = 0;  
        var j = 0;
        while(k < lenthbatchInput){  // varre cada coluna da linha selecionada
          
          var textoColuna = "";   // escreve espaço nas células preenchidas
          
          var dadosbatchInput = PagBatch.getRange(inicioLinhabatchInput,inicioColunabatchInput).setValue(textoColuna); // escreve espaço em branco na celula selecionada;
       
          inicioColunabatchInput++;        
          k++;
          
        }
        inicioLinhabatchInput++; 
  
      }
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao Limpar dados na Página Batch Input. " + erro + ". Contate o Administrador." );
  
  }   
  
  
}
