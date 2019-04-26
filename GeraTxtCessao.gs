function GeraTxtCessao() {
  
   /*
    Criação do arquivo Google Docs automaticamente dentro do Google Drive para inserção dos dados da Batch;
    Instanciar as planilhas Valores e Layout mostrando o tamanho de linha e coluna de cada planilha;
    Faz a varredura de linha a linha e coluna da planilha Valores, juntamente com a varedura de toda linha e coluna da planilha Layout.
    Ou seja, para cada linha lida da planilha Valores e´realizada uma varedora de todas as linhas e colunas da planilha Layuout e montado a Batch Input 
    no arquivo Google Docs com o nome  da Tela do SAP + a data do dia. 
    Funcition utilizada: 
                        RegraDeNegocio, PageLayoutCessao, PageValores,BatchMaster
 
 */
  
  try{
       // formata a data dd/mm/yyyy  
       var data = new Date(),
           dia  = data.getDate().toString(),
           diaF = (dia.length == 1) ? '0'+dia : dia,
           mes  = (data.getMonth()+1).toString(), //+1 pois no getMonth Janeiro começa com zero.
           mesF = (mes.length == 1) ? '0'+mes : mes,
           anoF = data.getFullYear();
  
        var dataHj = diaF+mesF+anoF; //  Gera a Data do dia
    
        // Vai na Planilha Layout e pega o conteúdo da celula C3 para gerar o nome do arquivo.
        var pagLayoutCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Cessão');
        var nomeArquivo = pagLayoutCessao.getRange('C2').getValue();
        
        Browser.msgBox( "Transação utilizada é: " + nomeArquivo  );
    
        var arquivo = nomeArquivo + "Cessao_" + dataHj;
   
        
        var arquivoFinal =  DocumentApp.create(arquivo); 
        //Obtem o ID do arquivo criado       
        var fb01_ID = arquivoFinal.getId();
    
    Browser.msgBox( "Arquivo do dia é: " + arquivo );  
  
  }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao criar o arquivo. Contate o Administrador." );
  
  }      

  // instanciar as Planilhas (Valores e Layout)
  var PagBatchInputCessao= new PageBatchInputCessao();
  var PlanLayoutCessao = new PageLayoutCessao();
  var ValoresLinha = 5;
  var linhatxt = 2;
  var ResultadoLista = [];
  var ResultadoValores = [];
  ResultadoValores[ValoresLinha] = "";
    
  
  var i = 0;
  var contToSave = 0;
  var contToSaveOder = 0;
  try{
    for ( i in PagBatchInputCessao){
          
          var colunaValores = PagBatchInputCessao[i];          
          var j = 0;
          
          var layoutLinha = 5;  // variável que informa a linha de inicio da planilha layout
          
          for(j in PlanLayoutCessao){
            
            var colunaLayout = PlanLayoutCessao[j];
            var k = 0;
      
              //Anda em cada coluna dessa linha
              if(k == 0){    // Trata a coluna Tela
                //Trata campo Tela
                var CampoTela = colunaLayout[k].toString();
                var lengthTela = CampoTela.length;
                var pagLayoutCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Cessão');
                var TamanhoCaracterTela = pagLayoutCessao.getRange('A3').getValue(); 
                var DiferencaTela = TamanhoCaracterTela - lengthTela;
                var caracter = 0;
                var addEspaco = " ";
                var addCaracter = " ";
                var cont = 0;
                
                while(caracter < DiferencaTela){               
                  
                  CampoTela = CampoTela + addEspaco;
                  
                  caracter ++
                    cont ++; 
                  
                }
                
                var resultTela = CampoTela;
                
                k++;
                
              } // termina o if Tela
            
              if(k == 1){     // Trata a coluna Tela Tipo
              //Trata Tela Tipo
              var tipoTela = colunaLayout[k].toString();;
              var lengthTipoTela = tipoTela.length;
              var pagLayoutCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Cessão');
              var TamanhoCaracterTela = pagLayoutCessao.getRange('B3').getValue(); 
              var DiferencaTela = TamanhoCaracterTela - lengthTipoTela;
              var caracter = 2;
              var addEspaco = " ";
              var addCaracter = " ";
              var cont = 0;
              
                
             
              if(lengthTipoTela = TamanhoCaracterTela)
              {
               
                   var resultTipoTela = tipoTela.substring(0,TamanhoCaracterTela); 
              
              }
               
              if(lengthTipoTela < TamanhoCaracterTela){  //Entende-se que o LengthTela tenha 4 posições e o usuário esqueceu do colocar o X, o sistema trata isso...
               
                 var resultTipoTela = tipoTela + "X";
                
              } 
             
              k++;
               
            } // termina o if Tipo Tela
     
            if(k == 2){   // Trata a coluna Campo
              //Trata o Campo
              var campo = colunaLayout[k].toString();
              var lengthTelaCampo = campo.length;
              var pagLayoutCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Cessão');
              var TamanhoCaracterTela = pagLayoutCessao.getRange('C3').getValue(); 
              var DiferencaTelaCampo = TamanhoCaracterTela - lengthTelaCampo;
              var caracter = 0;
              var addEspaco = " ";
              var addCaracter = " ";
              var cont = 0;
              
              if(lengthTelaCampo = TamanhoCaracterTela)
              {
                while(caracter < DiferencaTelaCampo){               
                     
                     campo = campo + addEspaco;
                     
                     caracter ++
                       cont ++; 
                }
                var resultCampo = campo.substring(0,TamanhoCaracterTela);
              }
               
              if(lengthTelaCampo < TamanhoCaracterTela){  //Entende-se que o LengthTela tenha 4 posições e o usuário esqueceu do colocar o X, o sistema trata isso...
               
                   while(caracter < DiferencaTelaCampo){               
                     
                     campo = campo + addEspaco;
                     
                     caracter ++
                       cont ++; 
                     
                   }
                
                    var resultCampo = (campo.substring(0,TamanhoCaracterTela));
                
                
              } 
             
              k++;
              
            } // termina o if do Campo
 
            if(k == 3){  // trata a coluna Valor
              //Trata Coluna Valor
              var colunaValor = colunaLayout[k].toString();
              var lengthTelaColunaVa = colunaValor.length;
              var pagLayoutCessao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Cessão');
              var TamanhoCaracterTela = pagLayoutCessao.getRange('D3').getValue(); 
              var DiferencaTela = TamanhoCaracterTela - lengthTelaColunaVa;
              var caracter = 2;
              var addEspaco = " ";
              var addCaracter = " ";
              var cont = 0;
  
              var tipoVariavel = pagLayoutCessao.getRange('E'+ layoutLinha).getValue();  // sempre busca o E da linha do Layout 
              var trataTipo = pagLayoutCessao.getRange('E2').getValue();
             
              /* Condição que trata o tipo de variável, quando a Coluna En for iqual a Variavel-E2, a condição deverá ir na planilha Valores e buscar 
              * a celula que a COLUNA/VALORES indicar, ou seja ir na planilha Valores e trazer nessa celula o texto contido nessa celula.              
              */
                           
              if(tipoVariavel == trataTipo ){
                
                var colunaValor = (colunaLayout[k]).toString();
                
                var pagValores = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Batch Input Cessão');
                var celulaValores = colunaValor + ValoresLinha;
                var DadosCelulaValores = ((pagValores.getRange(celulaValores).getValue()).toString()).replace("/",".").replace("/","."); // faz o tratamento de / por .
                var lengtDadosValoes = DadosCelulaValores.length;
                var TamanhoCaracterTela = pagLayoutCessao.getRange('D3').getValue(); 
                var DiferencaTela = TamanhoCaracterTela - lengtDadosValoes;
                var caracter = 2;
                var addEspaco = " ";
                var addCaracter = " ";
                var cont = 0;
                
                
                   var resultColunaValor = DadosCelulaValores.substring(0,TamanhoCaracterTela);  //substrig pega o tamanho desejado do texto. ex. um texto de 60 caracteres, eu quero apenas as 50 posições
        
              }else{  // pega a primeira opção do trata a coluna/valor
              
                       var resultColunaValor = colunaValor.substring(0,TamanhoCaracterTela);
              
                   }      
              
              }// termina o if da coluna Valor
            
            var resultado = resultTela + resultTipoTela + resultCampo + resultColunaValor;
            
            var doc = DocumentApp.openById(fb01_ID);
            var inserirResultadoNoDocs = doc.getBody().appendListItem(resultado);
        

            layoutLinha ++;
            linhatxt++;
     
            
        } //termina o for do Layout

          ValoresLinha ++
          doc.saveAndClose();
     
      }//termina o For que varre a linha da Planilha Valores
        
        Browser.msgBox( "TXT Cessão criado com sucesso. " );

   }catch(erro){
         
            Browser.msgBox( "Ocorreu um erro ao gerar o TXT. " + erro + ". Contate o Administrador." );
  
   }      
 
 
}
