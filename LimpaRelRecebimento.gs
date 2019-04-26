function LimpaRelRecebimento() {
  
  var PaginaRecebimento = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatorio Recebimento');
  var inicioLinhaRecebimento = 2;  
  var inicioColunaRecebimento = 1;
  
  
  // instancia as planilhas
  var recebimento = new PageRelRecebimento();
  
  try{
    
     //funciona também
  
      for(i in recebimento){
       
        var inicioColunaRecebimento = 1;
        var colunaRecebimento = recebimento[i];
        var lenthRecebimento = colunaRecebimento.length;
        var k = 0;  
        var j = 0;
        while(k < lenthRecebimento){  // varre cada coluna da linha selecionada
          
          var textoColuna = "";   // escreve espaço nas células preenchidas
          
          var dadosRecebimento = PaginaRecebimento.getRange(inicioLinhaRecebimento,inicioColunaRecebimento).setValue(textoColuna); // escreve espaço em branco na celula selecionada;
       
          inicioColunaRecebimento++;        
          k++;
          
        }
        inicioLinhaRecebimento++; 
  
      }
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao Limpar dados na Página Limpa Relatório Recebimento. " + erro + ". Contate o Administrador." );
  
  }   
  
  
}
