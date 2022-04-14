function cadastrarCabecario() {

  //Buscar informações

  var ssFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FORMULARIO');
  var dados = ssFormulario.getRange('b3:b7').getValues();
  var dadosFinais = [dados];
  var dadosDocumento = ssFormulario.getRange('e2').getValue();
  var dadosData = ssFormulario.getRange('a2').getValue();
  var dadosQtdeTotal = ssFormulario.getRange('c14').getValue();
  var dadosSubtotal = ssFormulario.getRange('e14').getValue();
  var dadosDesconto = ssFormulario.getRange('e15').getValue();
  var dadosTotal = ssFormulario.getRange('e16').getValue();
  var dadosFormapagto = ssFormulario.getRange('d17').getValue();
  var dadosPagto = ssFormulario.getRange('e17').getValue();
  var dadosTroco = ssFormulario.getRange('e18').getValue();
  var dadosVendedor = ssFormulario.getRange('b15').getValue();
  
  var dadosReferencia1 = ssFormulario.getRange('a9').getValue();
  var dadosDescricao1 = ssFormulario.getRange('b9').getValue();
  var dadosQtde1 = ssFormulario.getRange('c9').getValue();
  var dadosVUnit1 = ssFormulario.getRange('d9').getValue();
  var dadosTotal1 = ssFormulario.getRange('e9').getValue();

  var dadosReferencia2 = ssFormulario.getRange('a10').getValue();
  var dadosDescricao2 = ssFormulario.getRange('b10').getValue();
  var dadosQtde2 = ssFormulario.getRange('c10').getValue();
  var dadosVUnit2 = ssFormulario.getRange('d10').getValue();
  var dadosTotal2 = ssFormulario.getRange('e10').getValue();

  var dadosReferencia3 = ssFormulario.getRange('a11').getValue();
  var dadosDescricao3 = ssFormulario.getRange('b11').getValue();
  var dadosQtde3 = ssFormulario.getRange('c11').getValue();
  var dadosVUnit3 = ssFormulario.getRange('d11').getValue();
  var dadosTotal3 = ssFormulario.getRange('e11').getValue();

  var dadosReferencia4 = ssFormulario.getRange('a12').getValue();
  var dadosDescricao4 = ssFormulario.getRange('b12').getValue();
  var dadosQtde4 = ssFormulario.getRange('c12').getValue();
  var dadosVUnit4 = ssFormulario.getRange('d12').getValue();
  var dadosTotal4 = ssFormulario.getRange('e12').getValue();

  var dadosReferencia5 = ssFormulario.getRange('a13').getValue();
  var dadosDescricao5 = ssFormulario.getRange('b13').getValue();
  var dadosQtde5 = ssFormulario.getRange('c13').getValue();
  var dadosVUnit5 = ssFormulario.getRange('d13').getValue();
  var dadosTotal5 = ssFormulario.getRange('e13').getValue();

  //Verificar dados obrigatórios || significa 'ou' 
  //if (dados [0] == "" || dados[2]=="" || dados[5]=="") {
   // SpreadsheetApp.getUi().alert('Falta preencher dados obrigatórios!');
    //return;
  //}

  //Pegar aba banco de dados e buscar linha alvo na planilha CABECARIO_VENDAS
  var ssCabecarioVendas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CABECARIO_VENDAS');
  var ultimaLinha = ssCabecarioVendas.getLastRow();
  var linhaAlvo = ultimaLinha + 1;

  //Pegar aba banco de dados e buscar linha alvo na planilha ITENS_VENDA
  var ssItensVenda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ITENS_VENDA');
  var ultimaLinhaItensVenda = ssItensVenda.getLastRow();
  var linhaAlvoItensVenda = ultimaLinhaItensVenda + 1;

  //Teste nome duplicado
  //var nomesAtuais = ssBancoFuncionarios.getRange(1,1,ultimaLinha,1).getValues();

  //for (let i=0;i<nomesAtuais.length;i++){
   // var nomeAtual = String(nomesAtuais[i]).trim().toLowerCase();
   // var nomeTeste = String(dados[0]).trim().toLowerCase();

    //if (nomeAtual == nomeTeste) {
      //SpreadsheetApp.getUi().alert('Funcionário já cadastrado');
      //return;
    //}
  //}
  
  //Escrever dados finais

  ssCabecarioVendas.getRange(linhaAlvo,1,1,1).setValue(dadosDocumento);
  ssCabecarioVendas.getRange(linhaAlvo,2,1,1).setValue(dadosData);  
  ssCabecarioVendas.getRange(linhaAlvo,3,1,dados.length).setValues(dadosFinais);
  ssCabecarioVendas.getRange(linhaAlvo,8,1,1).setValue(dadosQtdeTotal); 
  ssCabecarioVendas.getRange(linhaAlvo,9,1,1).setValue(dadosSubtotal); 
  ssCabecarioVendas.getRange(linhaAlvo,10,1,1).setValue(dadosDesconto); 
  ssCabecarioVendas.getRange(linhaAlvo,11,1,1).setValue(dadosTotal); 
  ssCabecarioVendas.getRange(linhaAlvo,12,1,1).setValue(dadosFormapagto); 
  ssCabecarioVendas.getRange(linhaAlvo,13,1,1).setValue(dadosPagto);
  ssCabecarioVendas.getRange(linhaAlvo,14,1,1).setValue(dadosTroco);
  ssCabecarioVendas.getRange(linhaAlvo,15,1,1).setValue(dadosVendedor);


  ssItensVenda.getRange(linhaAlvoItensVenda,1,1,1).setValue(dadosDocumento);
  ssItensVenda.getRange(linhaAlvoItensVenda,2,1,1).setValue(dadosReferencia1);
  ssItensVenda.getRange(linhaAlvoItensVenda,3,1,1).setValue(dadosDescricao1);
  ssItensVenda.getRange(linhaAlvoItensVenda,4,1,1).setValue(dadosQtde1);
  ssItensVenda.getRange(linhaAlvoItensVenda,5,1,1).setValue(dadosVUnit1);
  ssItensVenda.getRange(linhaAlvoItensVenda,6,1,1).setValue(dadosTotal1);

  if (dadosReferencia2 != ""){
  
  linhaAlvoItensVenda=linhaAlvoItensVenda+1
  ssItensVenda.getRange(linhaAlvoItensVenda,1,1,1).setValue(dadosDocumento);
  ssItensVenda.getRange(linhaAlvoItensVenda,2,1,1).setValue(dadosReferencia2);
  ssItensVenda.getRange(linhaAlvoItensVenda,3,1,1).setValue(dadosDescricao2);
  ssItensVenda.getRange(linhaAlvoItensVenda,4,1,1).setValue(dadosQtde2);
  ssItensVenda.getRange(linhaAlvoItensVenda,5,1,1).setValue(dadosVUnit2);
  ssItensVenda.getRange(linhaAlvoItensVenda,6,1,1).setValue(dadosTotal2);

  }

  if (dadosReferencia3 != ""){
  
  linhaAlvoItensVenda=linhaAlvoItensVenda+1
  ssItensVenda.getRange(linhaAlvoItensVenda,1,1,1).setValue(dadosDocumento);
  ssItensVenda.getRange(linhaAlvoItensVenda,2,1,1).setValue(dadosReferencia3);
  ssItensVenda.getRange(linhaAlvoItensVenda,3,1,1).setValue(dadosDescricao3);
  ssItensVenda.getRange(linhaAlvoItensVenda,4,1,1).setValue(dadosQtde3);
  ssItensVenda.getRange(linhaAlvoItensVenda,5,1,1).setValue(dadosVUnit3);
  ssItensVenda.getRange(linhaAlvoItensVenda,6,1,1).setValue(dadosTotal3);

  }

  if (dadosReferencia4 != ""){
  
  linhaAlvoItensVenda=linhaAlvoItensVenda+1
  ssItensVenda.getRange(linhaAlvoItensVenda,1,1,1).setValue(dadosDocumento);
  ssItensVenda.getRange(linhaAlvoItensVenda,2,1,1).setValue(dadosReferencia4);
  ssItensVenda.getRange(linhaAlvoItensVenda,3,1,1).setValue(dadosDescricao4);
  ssItensVenda.getRange(linhaAlvoItensVenda,4,1,1).setValue(dadosQtde4);
  ssItensVenda.getRange(linhaAlvoItensVenda,5,1,1).setValue(dadosVUnit4);
  ssItensVenda.getRange(linhaAlvoItensVenda,6,1,1).setValue(dadosTotal4);

  }

  if (dadosReferencia5 != ""){
  
  linhaAlvoItensVenda=linhaAlvoItensVenda+1
  ssItensVenda.getRange(linhaAlvoItensVenda,1,1,1).setValue(dadosDocumento);
  ssItensVenda.getRange(linhaAlvoItensVenda,2,1,1).setValue(dadosReferencia5);
  ssItensVenda.getRange(linhaAlvoItensVenda,3,1,1).setValue(dadosDescricao5);
  ssItensVenda.getRange(linhaAlvoItensVenda,4,1,1).setValue(dadosQtde5);
  ssItensVenda.getRange(linhaAlvoItensVenda,5,1,1).setValue(dadosVUnit5);
  ssItensVenda.getRange(linhaAlvoItensVenda,6,1,1).setValue(dadosTotal5);

  }
  //SpreadsheetApp.getUi().alert(dados);
  // ctrl + s => para salvar;

  // Apagar os dados do formulário
  //ssFormulario.getRange('B3').clearContent();
  //ssFormulario.getRange('B4').clearContent();
  //ssFormulario.getRange('B5').clearContent();
  //ssFormulario.getRange('B6').clearContent();
  //ssFormulario.getRange('B7').clearContent();
 
 // limparDados()

  //ATUALIZAR O NUMERO DO DOCUMENTO EM +1
  var dadosDocumentoAtual=dadosDocumento+1
  ssFormulario.getRange('e2').setValue(dadosDocumentoAtual)

}

  function buscar(){

    //var ssFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FORMULARIO');
    var ss = SpreadsheetApp.getActive();
    var nome = ss.getRange('C4').getValue();
    var sslista = ss.getSheetByName('BANCO DE FUNCIONÁRIOS');

    if (nome == ""){
      //SpreadsheetApp.getUi().alert('Favor selecionar um funcionário');
      return false;

    }

    var ult_func = sslista.getLastRow();
    var funcdados = sslista.getRange(2,1,ult_func-1,6).getValues();

    var dados = [];
    for (let i=0; i<funcdados.length;i++){
      if (funcdados[i][0]==nome){
        for (let y=0; y<6; y++){
          dados.push(funcdados[i][y]);
        }

      break
      }
    }

    var ss = ss.getSheetByName('FORMULARIO');
    
    for (let i=0; i<9; i++){
      ss.getRange(i+6,3,1,1).setValue(dados[i]);
    }

    //SpreadsheetApp.getUi().alert(dados);

    //Segue abaixo outra maneira de preencher sem ser com o for
    //ss.getRange('c6').setValue(dados[0]);
    //ss.getRange('c7').setValue(dados[1]);
    //ss.getRange('c8').setValue(dados[2]);
    //ss.getRange('c9').setValue(dados[3]);
    //ss.getRange('c10').setValue(dados[4]);
    //ss.getRange('c11').setValue(dados[5]);

  }


function onEditAction(e){
  //buscar aba e intervalo
  var ss = e.source.getActiveSheet();
  var range = e.range.getA1Notation();

  //Verificar se a celula alvo foi editada
  if(ss.getName() != 'FORMULARIO' || range != 'C20'){
    return
  }

//pegar a função a ser utilizada

var acao = ss.getRange('B20').getValue();
if (acao == 'Finalizar Venda'){
  cadastrarCabecario();
} else if(acao == 'Buscar') {
  buscar();
} else {
  limparDados()
}

ss.getRange('C20').clearContent()
  
//SpreadsheetApp.getUi().alert('modificado')

}

function limparDados(){

var ss = SpreadsheetApp.getActive();
   // Apagar os dados do formulário
  ss.getRange('B3').clearContent();
  ss.getRange('B4').clearContent();
  ss.getRange('B5').clearContent();
  ss.getRange('B6').clearContent();
  ss.getRange('B7').clearContent();

  ss.getRange('A9').clearContent();
  ss.getRange('A10').clearContent();
  ss.getRange('A11').clearContent();
  ss.getRange('A12').clearContent();
  ss.getRange('A13').clearContent();

  ss.getRange('B9').clearContent();
  ss.getRange('B10').clearContent();
  ss.getRange('B11').clearContent();
  ss.getRange('B12').clearContent();
  ss.getRange('B13').clearContent();

 // ss.getRange('C9').clearContent();
  ss.getRange('C9').setValue('1');
  ss.getRange('C10').clearContent();
  ss.getRange('C11').clearContent();
  ss.getRange('C12').clearContent();
  ss.getRange('C13').clearContent();

  ss.getRange('D9').clearContent();
  ss.getRange('D10').clearContent();
  ss.getRange('D11').clearContent();
  ss.getRange('D12').clearContent();
  ss.getRange('D13').clearContent();

  ss.getRange('E15').clearContent();
  ss.getRange('E17').clearContent();

}




function cadastrarItenCupom() {

  //Buscar informações
  var ssCupom=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CUPOM');
  var ssItens_Cupom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens_Cupom');
  var dadosRef = ssItens_Cupom.getRange('c3').getValue();
  var dadosQtde = ssItens_Cupom.getRange('c4').getValue();
  var dadosDescr = ssItens_Cupom.getRange('c5').getValue();
  var dadosValor = ssItens_Cupom.getRange('c6').getValue();

  
  
  //Escrever dados finais

  ssCupom.getRange(11,1,1,1).setValue(dadosRef);
  ssCupom.getRange(11,2,1,1).setValue(dadosDescr);
  ssCupom.getRange(12,1,1,1).setValue(dadosQtde);
  ssCupom.getRange(12,2,1,1).setValue(dadosValor);
  

 // limparDados()

  //ATUALIZAR O NUMERO DO DOCUMENTO EM +1
  //var dadosDocumentoAtual=dadosDocumento+1
  //ssFormulario.getRange('e2').setValue(dadosDocumentoAtual)

}




function gerarCupom() {

  //Buscar informações

  var ssFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FORMULARIO');
  var dados = ssFormulario.getRange('b3:b7').getValues();
  var dadosFinais = [dados];
  var dadosDocumento = ssFormulario.getRange('e2').getValue();
  var dadosData = ssFormulario.getRange('a2').getValue();
  var dadosQtdeTotal = ssFormulario.getRange('c14').getValue();
  var dadosSubtotal = ssFormulario.getRange('e14').getValue();
  var dadosDesconto = ssFormulario.getRange('e15').getValue();
  var dadosTotal = ssFormulario.getRange('e16').getValue();
  var dadosFormapagto = ssFormulario.getRange('d17').getValue();
  var dadosPagto = ssFormulario.getRange('e17').getValue();
  var dadosTroco = ssFormulario.getRange('e18').getValue();
  var dadosVendedor = ssFormulario.getRange('b15').getValue();
  
  var dadosReferencia1 = ssFormulario.getRange('a9').getValue();
  var dadosDescricao1 = ssFormulario.getRange('b9').getValue();
  var dadosQtde1 = ssFormulario.getRange('c9').getValue();
  var dadosVUnit1 = ssFormulario.getRange('d9').getValue();
  var dadosTotal1 = ssFormulario.getRange('e9').getValue();

  var dadosReferencia2 = ssFormulario.getRange('a10').getValue();
  var dadosDescricao2 = ssFormulario.getRange('b10').getValue();
  var dadosQtde2 = ssFormulario.getRange('c10').getValue();
  var dadosVUnit2 = ssFormulario.getRange('d10').getValue();
  var dadosTotal2 = ssFormulario.getRange('e10').getValue();

  var dadosReferencia3 = ssFormulario.getRange('a11').getValue();
  var dadosDescricao3 = ssFormulario.getRange('b11').getValue();
  var dadosQtde3 = ssFormulario.getRange('c11').getValue();
  var dadosVUnit3 = ssFormulario.getRange('d11').getValue();
  var dadosTotal3 = ssFormulario.getRange('e11').getValue();

  var dadosReferencia4 = ssFormulario.getRange('a12').getValue();
  var dadosDescricao4 = ssFormulario.getRange('b12').getValue();
  var dadosQtde4 = ssFormulario.getRange('c12').getValue();
  var dadosVUnit4 = ssFormulario.getRange('d12').getValue();
  var dadosTotal4 = ssFormulario.getRange('e12').getValue();

  var dadosReferencia5 = ssFormulario.getRange('a13').getValue();
  var dadosDescricao5 = ssFormulario.getRange('b13').getValue();
  var dadosQtde5 = ssFormulario.getRange('c13').getValue();
  var dadosVUnit5 = ssFormulario.getRange('d13').getValue();
  var dadosTotal5 = ssFormulario.getRange('e13').getValue();

  //Verificar dados obrigatórios || significa 'ou' 
  //if (dados [0] == "" || dados[2]=="" || dados[5]=="") {
   // SpreadsheetApp.getUi().alert('Falta preencher dados obrigatórios!');
    //return;
  //}

  //Pegar aba banco de dados e buscar linha alvo na planilha CUPOM
  var ssCupom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CUPOM');
  var ultimaLinhaCupom = ssCupom.getLastRow();
  var linhaAlvoCupom = ultimaLinhaCupom + 1;

  //Pegar aba banco de dados e buscar linha alvo na planilha CABECARIO_VENDAS
  var ssCabecarioVendas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CABECARIO_VENDAS');
  var ultimaLinha = ssCabecarioVendas.getLastRow();
  var linhaAlvo = ultimaLinha + 1;

  //Pegar aba banco de dados e buscar linha alvo na planilha ITENS_VENDA
  var ssItensVenda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ITENS_VENDA');
  var ultimaLinhaItensVenda = ssItensVenda.getLastRow();
  var linhaAlvoItensVenda = ultimaLinhaItensVenda + 1;

  //Teste nome duplicado
  //var nomesAtuais = ssBancoFuncionarios.getRange(1,1,ultimaLinha,1).getValues();

  //for (let i=0;i<nomesAtuais.length;i++){
   // var nomeAtual = String(nomesAtuais[i]).trim().toLowerCase();
   // var nomeTeste = String(dados[0]).trim().toLowerCase();

    //if (nomeAtual == nomeTeste) {
      //SpreadsheetApp.getUi().alert('Funcionário já cadastrado');
      //return;
    //}
  //}
  
  //Escrever dados finais

  //ssCabecarioVendas.getRange(linhaAlvo,1,1,1).setValue(dadosDocumento);
  //ssCabecarioVendas.getRange(linhaAlvo,2,1,1).setValue(dadosData);  
  //ssCabecarioVendas.getRange(linhaAlvo,3,1,dados.length).setValues(dadosFinais);
  //ssCabecarioVendas.getRange(linhaAlvo,8,1,1).setValue(dadosQtdeTotal); 
  //ssCabecarioVendas.getRange(linhaAlvo,9,1,1).setValue(dadosSubtotal); 
  //ssCabecarioVendas.getRange(linhaAlvo,10,1,1).setValue(dadosDesconto); 
  //ssCabecarioVendas.getRange(linhaAlvo,11,1,1).setValue(dadosTotal); 
  //ssCabecarioVendas.getRange(linhaAlvo,12,1,1).setValue(dadosFormapagto); 
  //ssCabecarioVendas.getRange(linhaAlvo,13,1,1).setValue(dadosPagto);
  //ssCabecarioVendas.getRange(linhaAlvo,14,1,1).setValue(dadosTroco);
  //ssCabecarioVendas.getRange(linhaAlvo,15,1,1).setValue(dadosVendedor);


  ssCupom.getRange(8,2,1,1).setValue(dadosDocumento);

  ssCupom.getRange(linhaAlvoCupom,1,1,1).setValue(dadosReferencia1);
  ssCupom.getRange(linhaAlvoCupom,2,1,1).setValue(dadosDescricao1);
  ssCupom.getRange(linhaAlvoCupom+1,1,1,1).setValue(dadosQtde1);
  ssCupom.getRange(linhaAlvoCupom+1,2,1,1).setValue(dadosVUnit1);
  ssCupom.getRange(linhaAlvoCupom+1,4,1,1).setValue(dadosTotal1);

  if (dadosReferencia2 != ""){
  
  linhaAlvoCupom=linhaAlvoCupom+2
  
  ssCupom.getRange(linhaAlvoCupom,1,1,1).setValue(dadosReferencia2);
  ssCupom.getRange(linhaAlvoCupom,2,1,1).setValue(dadosDescricao2);
  ssCupom.getRange(linhaAlvoCupom+1,1,1,1).setValue(dadosQtde2);
  ssCupom.getRange(linhaAlvoCupom+1,2,1,1).setValue(dadosVUnit2);
  ssCupom.getRange(linhaAlvoCupom+1,4,1,1).setValue(dadosTotal2);

  }

if (dadosReferencia3 != ""){
  
  linhaAlvoCupom=linhaAlvoCupom+2
  
  ssCupom.getRange(linhaAlvoCupom,1,1,1).setValue(dadosReferencia3);
  ssCupom.getRange(linhaAlvoCupom,2,1,1).setValue(dadosDescricao3);
  ssCupom.getRange(linhaAlvoCupom+1,1,1,1).setValue(dadosQtde3);
  ssCupom.getRange(linhaAlvoCupom+1,2,1,1).setValue(dadosVUnit3);
  ssCupom.getRange(linhaAlvoCupom+1,4,1,1).setValue(dadosTotal3);

  }
  if (dadosReferencia4 != ""){
  
  linhaAlvoCupom=linhaAlvoCupom+2
  
  ssCupom.getRange(linhaAlvoCupom,1,1,1).setValue(dadosReferencia4);
  ssCupom.getRange(linhaAlvoCupom,2,1,1).setValue(dadosDescricao4);
  ssCupom.getRange(linhaAlvoCupom+1,1,1,1).setValue(dadosQtde4);
  ssCupom.getRange(linhaAlvoCupom+1,2,1,1).setValue(dadosVUnit4);
  ssCupom.getRange(linhaAlvoCupom+1,4,1,1).setValue(dadosTotal4);

  }

  if (dadosReferencia5 != ""){
  
  linhaAlvoCupom=linhaAlvoCupom+2 
 
  ssCupom.getRange(linhaAlvoCupom,1,1,1).setValue(dadosReferencia5);
  ssCupom.getRange(linhaAlvoCupom,2,1,1).setValue(dadosDescricao5);
  ssCupom.getRange(linhaAlvoCupom+1,1,1,1).setValue(dadosQtde5);
  ssCupom.getRange(linhaAlvoCupom+1,2,1,1).setValue(dadosVUnit5);
  ssCupom.getRange(linhaAlvoCupom+1,4,1,1).setValue(dadosTotal5);

  }

  //SpreadsheetApp.getUi().alert(dados);
  // ctrl + s => para salvar;

  // Apagar os dados do formulário
  //ssFormulario.getRange('B3').clearContent();
  //ssFormulario.getRange('B4').clearContent();
  //ssFormulario.getRange('B5').clearContent();
  //ssFormulario.getRange('B6').clearContent();
  //ssFormulario.getRange('B7').clearContent();
  limparDados()

  //ATUALIZAR O NUMERO DO DOCUMENTO EM +1
  // var dadosDocumentoAtual=dadosDocumento+1
  //ssFormulario.getRange('e2').setValue(dadosDocumentoAtual)

}

