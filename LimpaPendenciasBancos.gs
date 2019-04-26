function LimpaPendenciasBancos() {
  
  var PaginaPendenciaBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pendencias de Bancos');
  var inicioLinhaPendencia = 2;  
  var inicioColunaPendencia = 1;
  
  
  // instancia as planilhas
  var pendencia = new PagePendenciasBancos();
  
  try{
    
     //funciona também
  
      for(i in pendencia){
       
        var inicioColunaPendencia = 1;
        var colunapendencia = pendencia[i];
        var lenthpendencia = colunapendencia.length;
        var k = 0;  
        var j = 0;
        while(k < lenthpendencia){  // varre cada coluna da linha selecionada
          
          var textoColuna = "";   // escreve espaço nas células preenchidas
          
          var dadosPendencia = PaginaPendenciaBanco.getRange(inicioLinhaPendencia,inicioColunaPendencia).setValue(textoColuna); // escreve espaço em branco na celula selecionada;
       
          inicioColunaPendencia++;        
          k++;
          
        }
        inicioLinhaPendencia++; 
  
      }
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao Limpar dados na Página Pendencias de Bancos. " + erro + ". Contate o Administrador." );
  
  }   
  
  
}
