//Converter o Relatório 301 em arquivo google planilha
async function converterExcelParaGoogleSheets() {
  let nome = 'SCSC301'
  var idPastaOrigem = ""; // id da pasta onde esta o arquivo xlsx
  var idPastaDestino = ""; // id da pasta onde irá salvar o arquivo google sheets
  var pastaDestino = DriveApp.getFolderById(idPastaDestino)
  var arquivos = pastaDestino.getFiles()
  var arquivo = arquivos.next()
  var id = arquivo.getId()
  Logger.log(id)
  if  (arquivo == nome) {
  DriveApp.getFolderById(id).setTrashed(true);
  }

  var files = DriveApp.getFolderById(idPastaOrigem).getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while(files.hasNext()){
    var file = files.next();
    var name = file.getName().split('.')[0]; 
    var id = file.getId();
    var blob = file.getBlob();
 
    var newFile = {
        title : name,
        parents: [{id: idPastaDestino}] 
      }; 
      
    var sheetFile = Drive.Files.insert(newFile, blob, { convert: true });
  }

  pastaDestino = DriveApp.getFolderById(idPastaDestino)
  arquivos = pastaDestino.getFiles()
  arquivo = arquivos.next()
  id = arquivo.getId()
  Logger.log(id)
  if  (arquivo == nome) {
  return id
  }

}

function indiceColuna(x, y, z) {
  //exemplo: indiceColuna("texto a procurar","na linha","na Planilha")
  let index = z.getDataRange().getValues()[y - 1].indexOf(x);
  return index + 1
}

async function pesquisarPedido() {
  var idPlan = await converterExcelParaGoogleSheets()
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

      let empresa = dados[linha][0]
      let sc = dados[linha][1]
      let seq = dados[linha][3]
      let produto = dados[linha][4]
      let descricao = dados[linha][5]
      let obs = dados[linha][8]
      let um = dados[linha][13]

      guiamenu.getRange(proxProd, indiceColuna('Empresa', 4, guiamenu)).setValue(empresa)
      guiamenu.getRange(proxProd, indiceColuna('Solicitação', 4, guiamenu)).setValue(sc)
      guiamenu.getRange(proxProd, indiceColuna('Seq', 4, guiamenu)).setValue(seq)
      guiamenu.getRange(proxProd, indiceColuna('Produto', 4, guiamenu)).setValue(produto)
      guiamenu.getRange(proxProd, indiceColuna('Descrição', 4, guiamenu)).setValue(descricao)
      guiamenu.getRange(proxProd, indiceColuna('Observação', 4, guiamenu)).setValue(obs)
      guiamenu.getRange(proxProd, indiceColuna('U.M', 4, guiamenu)).setValue(um)

      proxProd++

    } else {
      if (x == 1) { return; }
    }

  }

  Browser.msgBox("Solicitação não foi encontrada no relatório")

}

function registrarPedido() {

}
