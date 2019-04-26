function AtualizaBaixas() {
  
   /*
  ** Informa como será o cabeçalho para colocar na página Baixa e por consequênte pegar os dados de cada linha + os dados de cada linha da página Base para a pag Baixas  
  */
  
  var Base = new PageBase();
  var pagBaixas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Baixas');
  var pagBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE');
 
  /*Retorna a posição da última linha que possui conteúdo e posiciona na 1ª linha para inserção do próximo dado. */  
  var lastRow_Baixas = pagBaixas.getLastRow();  //última linha preenchida da página Baixas
  var lastColumn_Baixas = pagBaixas.getLastColumn(); // ultima coluna preenchida da página Baixas
  
  var inicioLinhaBase = 2;  //inicio da linha Relatório Recebimento
  var inicioColunaBase = 1; // inicio da coluna Relatório Recebimento
  
  var limpaBaixas = new LimpaBaixas();  // limpa a página Baixas
  
  var i=0;
  var inicioLinhaBaixas = 2; // inicializa a linha da Baixas
  var inicioColunaBaixas = 1;   // inicializar a coluna da Batch
  
  
  
  try{
    
        for(i in Base){  // varre toda a página Relatório Recebimento
          
            var CodEmp = pagBase.getRange('A' + inicioLinhaBase).getValue();  // Vai no campo A+i e pega a data do documento 
            var CodCtaRaz = pagBase.getRange('B' + inicioLinhaBase).getValue();  // Vai no campo B+i e pega o código da empresa 
            var DataDoc = pagBase.getRange('F' + inicioLinhaBase).getValue();  // Vai no campo F+i e pega a conta razão
            var ObsHistorico = pagBase.getRange('H' + inicioLinhaBase).getValue();  // Vai no campo H+i e pega o código do cliente            
            var Valor = pagBase.getRange('I' + inicioLinhaBase).getValue();  // Vai no campo I+i e pega o valor R$ 
            var TipoDocumento = pagBase.getRange('D' + inicioLinhaBase).getValue();  // Vai no campo D+i e pega o histórico 
        
            var linhaBase = Base[i]; // entrar na linha Relatório Recebimento
            var lenthBase = linhaBase.length;
            var inicioColunaBaixas = 1;   // inicializar a coluna da Batch
            
            var k = 0;  
            var j = 0;
          
          if(TipoDocumento == "BB" && Valor < 0){  //Insere os dados na página Baixa
            
              while(k < lastColumn_Baixas){ // pecorre na coluna até chegar no ultimo campo preenchido da página da Baixas
                
                    if(k == 0){
                      
                        var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue(CodEmp); // acrescentar o próximo registro na campo informado da página Baixas 
                      
                    } 
                    if(k == 1){
                      
                      
                        if(CodCtaRaz == "18221248"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("CHRYSLER"); // acrescentar o próximo registro na campo informado da página Baixas 
                        
                        }
                        if(CodCtaRaz == "18201161" || CodCtaRaz == "18201175" || CodCtaRaz == "18201162" || CodCtaRaz == "18201177"){
                        
                            var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("FIAT"); // acrescentar o próximo registro na campo informado da página Baixas                       
                        
                        }
                        if(CodCtaRaz == "18221132" || CodCtaRaz == "18201908" || CodCtaRaz == "18201132" ){
                          
                              var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("JEEP"); // acrescentar o próximo registro na campo informado da página Baixas                       
                          
                        }
                      
                    }   
                    if(k == 2){
                      
                       if(CodCtaRaz == "18221248"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("33"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                        }
                      
                       if(CodCtaRaz == "18201161"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("1"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                       }
                      
                      if(CodCtaRaz == "18201175"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("237"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      if(CodCtaRaz == "18201177"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("33"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      
                      if(CodCtaRaz == "18221132"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("1"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      
                      if(CodCtaRaz == "18201908"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("237"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      if(CodCtaRaz == "18201132"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("33"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      
                      if(CodCtaRaz == "18201162"){
                          
                             var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("341"); // acrescentar o próximo registro na campo informado da página Baixas 
                             var banco = copiaParaBaixas.getValue();
                        
                      }
                      
                    } 
                
                    if(k == 3){
                      
                        var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue(CodCtaRaz); // acrescentar o próximo registro na campo informado da página Baixas  
                      
                    }
                
                    if(k == 4){
                      
                        if (banco == "1" && CodEmp =="G164" ){
                            
                              var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("804002678"); // FIAT - acrescentar o próximo registro na campo informado da página Baixas  
                          
                        }
                        
                        if (banco == "1" && CodEmp =="G654" ){
                              
                                var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("804002944"); // JEEP - acrescentar o próximo registro na campo informado da página Baixas  
                            
                        }
                      
                        if (banco == "33"){
                            
                              var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("804002776"); // FIAT e JEEP - acrescentar o próximo registro na campo informado da página Baixas  
                          
                        }
                      
                        if (banco == "237" ){
                              
                                var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("804002763"); // FIAT e JEEP - acrescentar o próximo registro na campo informado da página Baixas  
                            
                        }
                      
                        if (banco == "341" && CodEmp =="G164" ){
                            
                              var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue("804002774"); // FIAT - acrescentar o próximo registro na campo informado da página Baixas  
                          
                        }
                      
                    }
                
                    if(k == 5){
                      
                        var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue(DataDoc); // acrescentar o próximo registro na campo informado da página Baixas  
                      
                    }
                
                    if(k == 6){
                      
                          Valor = Valor * (-1); 
                          var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue(Valor); // acrescentar o próximo registro na campo informado da página Baixas
                   
                      
                    }
                
                    if(k == 7){
                      
                        var copiaParaBaixas = pagBaixas.getRange(inicioLinhaBaixas, inicioColunaBaixas).setValue(ObsHistorico); // acrescentar o próximo registro na campo informado da página Baixas  
                      
                    }
                    
                     
                  k ++;
                  inicioColunaBaixas ++; //pecorre na coluna Baixas
                
                
              }   
              
                inicioLinhaBaixas ++;   // pecorre na linha Baixas
               
            }
           inicioLinhaBase++;   // pecorre na linha da Página Base
        
        }
          Browser.msgBox( "Baixas atualizada com sucesso." );  
  
  }catch(erro){
    
    Browser.msgBox( "Ocorreu um erro ao inserir dados na Página Baixas. " + erro + ". Contate o Administrador." );  
  
  }
  
}
