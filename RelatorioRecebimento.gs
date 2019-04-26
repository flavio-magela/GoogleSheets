function RelatorioRecebimento() {
  
  var Baixas = new PageBaixas();   
  var RelatorioRecebimento = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatorio Recebimento'); 
  var PagBaixas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Baixas');
  var inicioLinhaRecebimento = 2;    
  var inicioLinhaBaixas = 2;
  
  var limpaRelRecebimento = new LimpaRelRecebimento();
  
  var i = 0;
  
  try{
    
          for(i in Baixas){
            
            var CodCliente = PagBaixas.getRange('E' + inicioLinhaBaixas).getValue();  // Vai no campo E+i e pega o valor 
            var CodBco = PagBaixas.getRange('C' + inicioLinhaBaixas).getValue();  // Vai no campo C+i e pega o valor 
            var ContaRazao = PagBaixas.getRange('D' + inicioLinhaBaixas).getValue();  // Vai no campo C+i e pega o valor 
            var inicioColunaRecebimento = 1;
            var inicoColunaExportaTxt = 1;
            
            var linhaPlanBaixas= Baixas[i];
            var lenthPlanBaixas = linhaPlanBaixas.length;
            var k = 0;  
            var j = 0;        
            
                while(k < lenthPlanBaixas){  // varre cada coluna da linha selecionada
                  
                  var textoColuna = linhaPlanBaixas[k]; // mostra o texto que está selecionado no momento                 
                  
                  
                  var dadosRecebimento = RelatorioRecebimento.getRange(inicioLinhaRecebimento,inicioColunaRecebimento).setValue(textoColuna);  // copia o texto para a celula selecionada 
                  
                  if (CodCliente == ""){                    
                    
                      if(CodBco == "1" && ContaRazao == "18201161"){
                        
                        textoColuna = "804002678";
                          
                        var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                      
                      }
                       if(CodBco == "1" && ContaRazao == "18221132"){
                        
                        textoColuna = "804002944";
                          
                        var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                      
                      }
                      if(CodBco == "237"){
                          
                          textoColuna = "804002763";
                            
                          var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                        
                       }
                       if(CodBco == "341"){
                          
                          textoColuna = "804002774";
                            
                          var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                        
                       }
                    
                       if(CodBco == "33" && ContaRazao == "18201177" || ContaRazao == "18201132"){
                          
                          textoColuna = "804002776";
                            
                          var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                        
                       } 
                      if(CodBco == "33" && ContaRazao == "18221248" ){
                          
                          textoColuna = "0001";
                            
                          var dadosRecebimento = RelatorioRecebimento.getRange('E' + inicioLinhaBaixas).setValue(textoColuna);
                        
                       } 
                    
                  }
               
                  inicioColunaRecebimento++;                     
                  k++;
                  
                }
             inicioLinhaRecebimento ++; 
             inicioLinhaBaixas ++;
                     
      
          }  
    
            //var PagePendenciasBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pendencias de Bancos');
            
            var GeraPendenciasBancos = new PendenciasBancos();   // Gera a Página Pendencias Bancos
            
            Browser.msgBox( "Base nova atualizada com sucesso." );
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao exportar os dados para a Pagina Relatório Recebimento. " + erro + ". Contate o Administrador." );
  
   }      
  
  
  
}

  

