function onEdit(e) {
  var sheet = e.source.getSheetByName('Sheet1');
  var pagosSheet = e.source.getSheetByName('Pagamentos');

  if (e.range.getSheet().getName() !== 'Sheet1') return;

  var row = e.range.getRow();
  var uniqueID = sheet.getRange(row, 3).getValue();
  var nomeCompleto = sheet.getRange(row, 4).getValue();
  var cpfCnpj = sheet.getRange(row, 7).getValue();
  var telefone = sheet.getRange(row, 11).getValue();
  var email = sheet.getRange(row, 12).getValue();
  var valorTotal = sheet.getRange(row, 13).getValue();
  var aVista = sheet.getRange(row, 14).getValue();
  var numParcelas = sheet.getRange(row, 15).getValue();
  var dataVencimento = sheet.getRange(row, 16).getValue();

  // Verifica se a linha já existe na aba "Pagamentos"
  var pagamentosData = pagosSheet.getRange('A:A').getValues();
  var exists = pagamentosData.some(function(row) { return row[0] == uniqueID; });

  if (!exists) {
    if (aVista === "Sim") {
      pagosSheet.appendRow([uniqueID, nomeCompleto, cpfCnpj, telefone, email, valorTotal, aVista, numParcelas, dataVencimento, 'Pago', new Date(), valorTotal, 1, 1]);
    } else {
      var parcelaValor = valorTotal / numParcelas;
      pagosSheet.appendRow([uniqueID, nomeCompleto, cpfCnpj, telefone, email, valorTotal, aVista, numParcelas, dataVencimento, 'Em Andamento', '', '', 0]);
      for (var i = 1; i <= numParcelas; i++) {
        pagosSheet.appendRow([uniqueID, nomeCompleto, cpfCnpj, telefone, email, parcelaValor, aVista, numParcelas, dataVencimento, 'Não Pago', '', parcelaValor, i, 0]);
      }
    }
  }
}

function recordPayment(uniqueID, paymentAmount, paymentDate, parcela) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pagosSheet = ss.getSheetByName('Pagamentos');
  var data = pagosSheet.getDataRange().getValues();

  var totalPaid = 0;
  var totalParcelas = 0;
  var numParcelas = 0;
  var totalValue = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == uniqueID) {
      var currentParcela = data[i][14];
      if (data[i][14] == parcela) {
        // Atualiza o valor pago e a data de pagamento
        pagosSheet.getRange(i + 1, 11).setValue(paymentDate); // Data de Pagamento
        pagosSheet.getRange(i + 1, 12).setValue(paymentAmount); // Valor Pago

        // Atualiza o total pago e o status
        totalPaid += paymentAmount;
        numParcelas = data[i][7];
        totalValue = data[i][5]; // Valor Total
        totalParcelas = pagosSheet.getRange('N:N').getValues().filter(row => row[0] == uniqueID).length;

        // Atualiza o número de parcelas pagas
        pagosSheet.getRange(i + 1, 13).setValue(totalParcelas);

        // Atualiza o status baseado no total pago
        if (totalPaid >= totalValue) {
          pagosSheet.getRange(i + 1, 10).setValue('Pago');
        } else {
          pagosSheet.getRange(i + 1, 10).setValue('Em Andamento');
        }
        break;
      }
    }
  }
}

function createTrigger() {
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}
