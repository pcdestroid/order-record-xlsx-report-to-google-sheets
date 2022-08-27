//Converter o Relatório 301 em arquivo google planilha
async function converterExcelParaGoogleSheets() {
  let nome = 'SCSC301'
  var idPastaOrigem = "1HX8E1CnWYyavLU6bkNQZaBWmSLT8YQH9";
  var idPastaDestino = "1znhgZ80zGTiJTLg0rHEJGhxUJzgPNqM8";
  var pastaDestino = DriveApp.getFolderById(idPastaDestino)
  var arquivos = pastaDestino.getFiles()
  var arquivo = arquivos.next()
  var id = arquivo.getId()
  Logger.log(id)
  if (arquivo == nome) {
    DriveApp.getFolderById(id).setTrashed(true);
  }

  var files = DriveApp.getFolderById(idPastaOrigem).getFilesByType(MimeType.MICROSOFT_EXCEL);

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName().split('.')[0];
    var id = file.getId();
    var blob = file.getBlob();

    var newFile = {
      title: name,
      parents: [{ id: idPastaDestino }]
    };

    var sheetFile = Drive.Files.insert(newFile, blob, { convert: true });
  }

  pastaDestino = DriveApp.getFolderById(idPastaDestino)
  arquivos = pastaDestino.getFiles()
  arquivo = arquivos.next()
  id = arquivo.getId()
  Logger.log(id)
  if (arquivo == nome) {
    return id
  }

}

function indiceColuna(x, y, z) {
  //exemplo: indiceColuna("texto a procurar","na linha","na Planilha")
  let index = z.getDataRange().getValues()[y - 1].indexOf(x);
  return index + 1
}

async function pesquisarPedido() {
  //var idPlan = await converterExcelParaGoogleSheets()
  var idPlan = '1KnYiT7oLS4KEcmD96UnANIxVgIvNR9fYtM5pcXRwu7k'
  var base = SpreadsheetApp.openById(idPlan)
  var guiaDados = base.getSheetByName('Plan1');
  var pedidos = SpreadsheetApp.getActiveSpreadsheet();
  var guiamenu = pedidos.getSheetByName('Registrar');
  var marina = guiamenu.getRange('B2').getValue();
  var solicitacao = guiamenu.getRange('D2').getValue();

  guiamenu.getRangeList(['A5:AB1000']).clear({ contentsOnly: true, skipFilteredRows: true });//Limpar lista

  if (marina == "") {
    Browser.msgBox("Precisa informar a Marina")
    return;
  }
  if (solicitacao == "") {
    Browser.msgBox("Precisa informar o número de solicitação")
    return;
  }

  var pesquisa = marina + solicitacao
  var dados = guiaDados.getRange(2, 2, guiaDados.getLastRow(), 45).getValues();
  var proxProd = 5
  var x = 0
  for (var linha = 0; linha < dados.length; linha++) {
    if (dados[linha][0] + dados[linha][1] == pesquisa) {
      x = 1
      
      let empresa = dados[linha][0],sc = dados[linha][1],seq = dados[linha][3],produto = dados[linha][4],descricao = dados[linha][5],obs = dados[linha][8],um = dados[linha][13];
      let familia = dados[linha][14],quantidade = dados[linha][15],dataSolicitacao = dados[linha][16],solicitanteSc = dados[linha][18],aprovadorSc = dados[linha][20],sitAprovSc = dados[linha][21];
      let dataAprovSc = dados[linha][22],nroOc = dados[linha][25],statusOc = dados[linha][26],seqOc = dados[linha][27],codForn = dados[linha][28],descricaoFornecedor = dados[linha][29];
      let qtdOc = dados[linha][31],valorOc = dados[linha][32],dataOc = dados[linha][33],usuarioGerouOc = dados[linha][36],sitAprovOc = dados[linha][38],aprovadorOc = dados[linha][41];
      let dataAprovOc = dados[linha][42]

      guiamenu.getRange(proxProd, indiceColuna('Empresa', 4, guiamenu)).setValue(empresa)
      guiamenu.getRange(proxProd, indiceColuna('Solicitação', 4, guiamenu)).setValue(sc)
      guiamenu.getRange(proxProd, indiceColuna('Seq', 4, guiamenu)).setValue(seq)
      guiamenu.getRange(proxProd, indiceColuna('Produto', 4, guiamenu)).setValue(produto)
      guiamenu.getRange(proxProd, indiceColuna('Descrição', 4, guiamenu)).setValue(descricao)
      guiamenu.getRange(proxProd, indiceColuna('Observação', 4, guiamenu)).setValue(obs)
      guiamenu.getRange(proxProd, indiceColuna('U.M', 4, guiamenu)).setValue(um)
      guiamenu.getRange(proxProd, indiceColuna('Família', 4, guiamenu)).setValue(familia)
      guiamenu.getRange(proxProd, indiceColuna('Quantidade', 4, guiamenu)).setValue(quantidade)
      guiamenu.getRange(proxProd, indiceColuna('Data Solic.', 4, guiamenu)).setValue(dataSolicitacao)
      guiamenu.getRange(proxProd, indiceColuna('Solicitante SC', 4, guiamenu)).setValue(solicitanteSc)
      guiamenu.getRange(proxProd, indiceColuna('Aprovador SC', 4, guiamenu)).setValue(aprovadorSc)
      guiamenu.getRange(proxProd, indiceColuna('Sit. Aprov.', 4, guiamenu)).setValue(sitAprovSc)
      guiamenu.getRange(proxProd, indiceColuna('Data  Aprov.', 4, guiamenu)).setValue(dataAprovSc)
      guiamenu.getRange(proxProd, indiceColuna('Nro OC', 4, guiamenu)).setValue(nroOc)
      guiamenu.getRange(proxProd, indiceColuna('Status da OC', 4, guiamenu)).setValue(statusOc)
      guiamenu.getRange(proxProd, indiceColuna('Seq. OC', 4, guiamenu)).setValue(seqOc)
      guiamenu.getRange(proxProd, indiceColuna('Cod. Fornec.', 4, guiamenu)).setValue(codForn)
      guiamenu.getRange(proxProd, indiceColuna('Descrição Fornecedor', 4, guiamenu)).setValue(descricaoFornecedor)
      guiamenu.getRange(proxProd, indiceColuna('Qtd. OC', 4, guiamenu)).setValue(qtdOc)
      guiamenu.getRange(proxProd, indiceColuna('Valor', 4, guiamenu)).setValue(valorOc)
      guiamenu.getRange(proxProd, indiceColuna('Data OC', 4, guiamenu)).setValue(dataOc)
      guiamenu.getRange(proxProd, indiceColuna('Usuário Gerou OC', 4, guiamenu)).setValue(usuarioGerouOc)
      guiamenu.getRange(proxProd, indiceColuna('Sit. Aprov. OC', 4, guiamenu)).setValue(sitAprovOc)
      guiamenu.getRange(proxProd, indiceColuna('Aprovador OC', 4, guiamenu)).setValue(aprovadorOc)
      guiamenu.getRange(proxProd, indiceColuna('Data Aprov. OC', 4, guiamenu)).setValue(dataAprovOc)

      proxProd++

    } else {
      if (x == 1) { return; }
    }

  }

  Browser.msgBox("Solicitação não foi encontrada no relatório")

}

function registrarPedido() {

}
