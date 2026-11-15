/**
 * Sistema de Governança Hospitalar • Pio XII
 * Backend API REST (Google Apps Script)
 * • Criado por 𝕄𝕣𝕝•𝕄𝕒𝕣𝕢𝕦𝕖𝕤 •
 */

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let aba = ss.getSheetByName("Auditorias");
    if (!aba) {
      aba = ss.insertSheet("Auditorias");
      aba.appendRow(["ID", "Data/Hora", "Auditora", "Setor", "Leito", "Modalidade", "Total Itens", "Não Conformes", "Falhas Detalhadas"]);
      aba.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#9A2B2E").setFontColor("#FFFFFF");
    }
    
    aba.appendRow([
      dados.id,
      dados.dataHora,
      dados.auditora,
      dados.setor,
      dados.leito,
      dados.tipoLimpeza,
      dados.totalItens,
      dados.naoConformes,
      dados.falhasStr
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "sucesso", id: dados.id }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (erro) {
    return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: erro.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("Auditorias");
    if (!aba) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }
    
    const range = aba.getDataRange().getValues();
    const cabecalho = range.shift();
    const historico = range.map(r => ({
      id: r[0],
      dataHora: r[1],
      auditora: r[2],
      setor: r[3],
      leito: r[4],
      tipoLimpeza: r[5],
      totalItens: r[6],
      naoConformes: r[7],
      falhasStr: r[8]
    })).reverse();
    
    return ContentService.createTextOutput(JSON.stringify(historico))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (erro) {
    return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: erro.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
