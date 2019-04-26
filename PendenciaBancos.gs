function PendenciasBancos() {
  
  var PagePendenciasBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pendencias de Bancos'); 
  var copiaParaPendenciaBaixa = PagePendenciasBanco.getRange(2,1).setValue("NÃO DEIXAR A PÁGINA EM BRANCO");// Caso a página fique em branco - Dá erro no proximo processamento da Página.
  var PagBaixas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Baixas');
  var Baixas = new PageBaixas();
  var PendenciasBancos = new PagePendenciasBancos()
  var inicioLinhaPendeciasBancos = 2;
  
  var inicioLinhaBaixas = 2;
  
  var i = 0;
  var j = 0;
  
 var LimpaPendenciaBanco = new LimpaPendenciasBancos();
  
  try{
    
    for (i in Baixas){
      
          var linhaValoresBaixas = Baixas[i];
          var inicioColunaPendeciasBancos  = 1;
          var lastLinhaBaixa = PagePendenciasBanco.getLastRow(); // mostra a última linha da Página Baixas
          var inicioLinhaBaixa = 2
          var IdentificaCampoBaixas = PagBaixas.getRange('E' + inicioLinhaBaixas).getValue();
          var linhaPendenciasBancos = PagePendenciasBanco.getRange(lastLinhaBaixa, inicioColunaPendeciasBancos).getValues(); // Mostra os dados da última linha da Página Baixas preenchida
          
          var k = 0;
          
          if (IdentificaCampoBaixas == ""){  //verifica se o campo E + i está vazio
            
                for(j in PendenciasBancos){   //varre a Página Pendências Banco até o final
                  
                  inicioLinhaPendeciasBancos ++;
                  
                }
                
                var lengthBaixas = linhaValoresBaixas.length;
            
                while(k < lengthBaixas){
                  
                  var campoBaixas = linhaValoresBaixas[k];                
                  //var copiaParaPendenciaBaixa = PagePendenciasBanco.getRange(lastLinhaBaixa +1, inicioColunaPendeciasBancos).setValue(campoBaixas); // acrescentar o próximo registro na 1ª linha vazia da página Pendencias Bancos (+1)
                  var copiaParaPendenciaBaixa = PagePendenciasBanco.getRange(inicioLinhaBaixa, inicioColunaPendeciasBancos).setValue(campoBaixas); // acrescentar no inicio da linha
                  k ++;
                  inicioColunaPendeciasBancos ++;
                  
                }               
            
          }
          
          inicioLinhaBaixas ++;
    }
    
  } catch(erro){
  
       Browser.msgBox( "Ocorreu um erro ao carregar as Pendências dos Bancos. " + erro + ". Contate o Administrador." );
  
  }
  
 
}
