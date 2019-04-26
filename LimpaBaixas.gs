function LimpaBaixas() {
  
  var PagBaixas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Baixas');
  var inicioLinhaBaixas = 2;  
  var inicioColunaBaixas = 1;
  
  
  // instancia as planilhas
  var baixas = new PageBaixas();
  
  try{
    
     //funciona também
  
      for(i in baixas){
       
        var inicioColunaBaixas = 1;
        var colunaBaixas = baixas[i];
        var lenthBaixas = colunaBaixas.length;
        var k = 0;  
        var j = 0;
        while(k < lenthBaixas){  // varre cada coluna da linha selecionada
          
          var textoColuna = "";   // escreve espaço nas células preenchidas
          
          var dadosBaixas = PagBaixas.getRange(inicioLinhaBaixas,inicioColunaBaixas).setValue(textoColuna); // escreve espaço em branco na celula selecionada;
       
          inicioColunaBaixas++;        
          k++;
          
        }
        inicioLinhaBaixas++; 
  
      }
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao Limpar dados na Página Limpa Baixas. " + erro + ". Contate o Administrador." );
  
  }   
  
  
}
