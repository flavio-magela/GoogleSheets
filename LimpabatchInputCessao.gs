function LimpabatchInputCessao() {
  
  var PagBatchCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input Cessão');
  var inicioLinhabatchInputCessao = 5;  
  var inicioColunabatchInputCessao = 2;
  
  
  // instancia as planilhas
  var batchInputCessao = new PageBatchInputCessao();
  
  try{
    
     //funciona também
  
      for(i in batchInputCessao){
       
        var inicioColunabatchInputCessao = 2;
        var colunabatchInputCessao = batchInputCessao[i];
        var lenthbatchInputCessao = colunabatchInputCessao.length;
        var k = 0;  
        var j = 0;
        while(k < lenthbatchInputCessao){  // varre cada coluna da linha selecionada
          
          var textoColuna = "";   // escreve espaço nas células preenchidas
          
          var dadosbatchInputCessao = PagBatchCessao.getRange(inicioLinhabatchInputCessao,inicioColunabatchInputCessao).setValue(textoColuna); // escreve espaço em branco na celula selecionada;
       
          inicioColunabatchInputCessao++;        
          k++;
          
        }
        inicioLinhabatchInputCessao++; 
  
      }
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao Limpar dados na Página Batch Input Cessão. " + erro + ". Contate o Administrador." );
  
  }   
  
  
}

