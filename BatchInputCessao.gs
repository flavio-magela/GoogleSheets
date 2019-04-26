function BatchInputCessao() {
  
  /*
  ** Informa como será o cabeçalho para colocar na página Batch Input e por consequênte pegar os dados de cada linha + os dados de cada linha da página Layout para gerar o .txt  
  */
  
  var recebimento = new PageRelRecebimento();
  //var batchInput = new PageBatchInput();
  var pagBatchInputCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input Cessão');
  var pagRecebimento = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatorio Recebimento');
 
  /*Retorna a posição da última linha que possui conteúdo e posiciona na 1ª linha para inserção do próximo dado. */  
  var lastRow_Batch = pagBatchInputCessao.getLastRow();  //última linha preenchida da página Batch Input
  var lastColumn_Batch = pagBatchInputCessao.getLastColumn(); // ultima coluna preenchida da página Batch Input
  
  var inicioLinhaRecebimento = 2;  //inicio da linha Relatório Recebimento
  var inicioColunaRecebimento = 1; // inicio da coluna Relatório Recebimento
  
  var limpaBatchCessao = new LimpabatchInputCessao();  // limpa a página Batch Input
  
  var i=0;
  var inicioLinhaBatch = 5; // inicializa a linha da Batch Input
  
  
  
  try{
    
        for(i in recebimento){  // varre toda a página Relatório Recebimento
          
            var DataDoc = pagRecebimento.getRange('F' + inicioLinhaRecebimento).getValue();  // Vai no campo F+i e pega a data do documento 
            var CodEmpresa = pagRecebimento.getRange('A' + inicioLinhaRecebimento).getValue();  // Vai no campo A+i e pega o código da empresa 
            var ContaRazao = pagRecebimento.getRange('D' + inicioLinhaRecebimento).getValue();  // Vai no campo D+i e pega a conta razão
            var CodCliente = pagRecebimento.getRange('E' + inicioLinhaRecebimento).getValue();  // Vai no campo E+i e pega o código do cliente            
            var Valor = pagRecebimento.getRange('G' + inicioLinhaRecebimento).getValue();  // Vai no campo G+i e pega o valor R$ 
            var ObsHistorico = pagRecebimento.getRange('H' + inicioLinhaRecebimento).getValue();  // Vai no campo H+i e pega o histórico 
          
          // Campos Fixo na página Batch Input   - Tada vez que é gerado a Batch Input esses campos são fixo.        
            var TipoDocFixo = "DZ";
            var ChaveBancoFixo = "40";
            var ChaveClienteFixo = "11";
            var ContaRazaoFixo = "M";
            
          
            var linhaRecebimento = recebimento[i]; // entrar na linha Relatório Recebimento
            var lenthRecebimento = linhaRecebimento.length;
            var inicioColunaBatch = 2;   // inicializar a coluna da Batch
            var k = 1;  
            var j = 0;
          
          if( CodCliente == "D_FATT0001" || CodCliente == "D_FATT0002" || CodCliente == "D_FATT0003" || CodCliente == "D_FATT0004" || CodCliente == "D_FATT0005" || CodCliente == "D_FATT0006"
             || CodCliente == "D_FATT0007" || CodCliente == "D_FATT0008" || CodCliente == "D_FATT0009" || CodCliente == "D_FATT0010" || CodCliente == "D_FATT0011" || CodCliente == "D_FATT0012" ||
             CodCliente == "D_FATT0013" || CodCliente == "D_FATT0014" || CodCliente == "D_FATT0015"){  
            
              while(k < lastColumn_Batch){ // pecorre na coluna até chegar no ultimo campo preenchido da página da Batch Input
                
                    if(k == 1){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(DataDoc); // acrescentar o próximo registro na campo informado da página Batch Input 
                      
                    } 
                    if(k == 2){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(TipoDocFixo); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }   
                    if(k == 3){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(CodEmpresa); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    } 
                    if(k == 4){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(DataDoc); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 5){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ChaveBancoFixo); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 6){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ContaRazao); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 7){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(Valor); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 8){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ObsHistorico); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 9){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ChaveClienteFixo); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 10){                  
                      
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(CodCliente); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 11){
                      
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ObsHistorico); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }
                    if(k == 12){
                                            
                        var copiaParaBatchInputCessao = pagBatchInputCessao.getRange(inicioLinhaBatch, inicioColunaBatch).setValue(ContaRazaoFixo); // acrescentar o próximo registro na campo informado da página Batch Input  
                      
                    }  
                     
                  k ++;
                  inicioColunaBatch ++; //pecorre na coluna Batch Input
                
              }   
              
                inicioLinhaBatch ++;   // pecorre na linha Batch Input
               
            }
           inicioLinhaRecebimento++;   // pecorre na linha Relatório Recebimento
        
        }
          Browser.msgBox( "Batch Input Cessão gerada com sucesso." );
  
  }catch(erro){
    
    Browser.msgBox( "Ocorreu um erro ao inserir dados na Página Batch Input Cessão. " + erro + ". Contate o Administrador." );  
  
  }
  
  
}

