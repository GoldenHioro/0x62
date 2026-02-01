function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AutoFill')
    .addItem('Gerar CoHo e Procuração', 'mostrarFormulario')
    .addToUi();
}

function mostrarFormulario() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt('Insira o Unique ID do cliente:');
  var clienteId = response.getResponseText().trim();
  
  if (!clienteId) {
    ui.alert('Nenhum ID de cliente fornecido.');
    return;
  }
  
  var valorResponse = ui.prompt('Insira o valor do contrato (R$). Ex: 1.000,00');
  var valor = valorResponse.getResponseText().trim();
  
  if (!valor) {
    ui.alert('Nenhum valor fornecido.');
    return;
  }

  var aVistaResponse = ui.alert('Pagamento à vista?', ui.ButtonSet.YES_NO);
  var aVista = (aVistaResponse == ui.Button.YES);

  var parcelas = null;
  if (!aVista) {
    var parcelasResponse = ui.prompt('Número de parcelas. Ex: 3');
    parcelas = parcelasResponse.getResponseText().trim();
    
    if (!parcelas) {
      ui.alert('Nenhum número de parcelas fornecido.');
      return;
    }
  }

  var diavencResponse = ui.prompt('Dia de vencimento das parcelas (1-31):');
  var diavenc_parcelas = diavencResponse.getResponseText().trim();
  
  processarFormulario(valor, aVista, parcelas, clienteId, diavenc_parcelas);
}

// Função que será chamada quando uma alteração é detectada na planilha
function onChange(e) {
  // Verifica se a mudança foi a adição de uma nova linha
  if (e.changeType === 'INSERT_ROW') {
    var sheet = e.source.getSheetByName('Dados');
    var lastRow = sheet.getLastRow();
    var newRow = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var clienteId = newRow[2]; // Supondo que o Unique ID esteja na coluna C (índice 2)
    if (clienteId) {
      processarFormulario(newRow[12], // valor
                          newRow[13] === 'Sim', // aVista
                          newRow[14], // parcelas
                          clienteId, // clienteId
                          newRow[15]); // diavenc_parcelas
    }
  }
}


function processarFormulario(valor, aVista, parcelas, clienteId, diavenc_parcelas) {
  valor = parseFloat(valor.replace(/\./g, '').replace(',', '.'));
  
  var valor_honocontrato = formatarValor(valor);
  var valor_honocontrato_extenso = converterParaExtenso(valor);

  var valorparcelado_honocontrato, valorparcelado_honocontrato_extenso, n_parcelas, texto_pagamento;

  if (aVista) {
    valorparcelado_honocontrato = valor_honocontrato;
    valorparcelado_honocontrato_extenso = valor_honocontrato_extenso;
    n_parcelas = 1;
    texto_pagamento = 'Pagamento à vista, efetivado no dia de assinatura do presente contrato.';
  } else {
    n_parcelas = parseInt(parcelas, 10);
    valorparcelado_honocontrato = formatarValor(valor / n_parcelas);
    valorparcelado_honocontrato_extenso = converterParaExtenso(valor / n_parcelas);
    texto_pagamento = `Dividido em ${n_parcelas} vezes, no valor de ${valorparcelado_honocontrato} (${valorparcelado_honocontrato_extenso}) cada parcela, com vencimento no ${diavenc_parcelas}º dia de cada mês.`;
  }

  var sheetId = '1UhTmg7okCm9sDApbWdTXCZ9JbiCbTuBdoHXtDI25HCg';
  var docTemplateId1 = '1GOHHcOusGM8wOfX2uZ0aOPsXoT6JYpySirfVdbu-0ho';
  var docTemplateId2 = '17XpXcKtlcS-8mZ9NE0A3kI2D7ULJKIvvFPLkhapeXPg';
  var rootFolderId = '1-7IKrkWNuRXc8YsyVc4OSEU5DECG7A7p';

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Dados');
  if (!sheet) {
    Logger.log('Aba "Dados" não encontrada.');
    return;
  }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var clienteDados = null;
  var idColumnIndex = 2;

  Logger.log('ID do cliente fornecido: "' + clienteId + '"');

  for (var i = 1; i < values.length; i++) {
    var uniqueId = String(values[i][idColumnIndex]).trim();

    if (uniqueId === clienteId) {
      clienteDados = values[i];
      break;
    }
  }

  if (clienteDados == null) {
    Logger.log("Cliente não encontrado.");
    return;
  }

  // Atualizar dados na planilha
  var valorTotalColumn = 13; // Coluna M
  var aVistaColumn = 14; // Coluna N
  var parcelasColumn = 15; // Coluna O
  var vencimentoColumn = 16; // Coluna P

  sheet.getRange(i + 1, valorTotalColumn).setValue(valor_honocontrato);
  sheet.getRange(i + 1, aVistaColumn).setValue(aVista ? 'Sim' : 'Não');
  sheet.getRange(i + 1, parcelasColumn).setValue(n_parcelas);
  sheet.getRange(i + 1, vencimentoColumn).setValue(diavenc_parcelas);

  function obterMesEmPortugues(mesNumero) {
    var meses = [
      'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
      'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ];
    return meses[mesNumero - 1];
  }

  function preencherEDatarDocumento(docTemplateId, clienteDados, nomeDocumento, folder) {
    var newDocName = nomeDocumento + ' - ' + clienteDados[3];
    var newDoc = DriveApp.getFileById(docTemplateId).makeCopy(newDocName);
    var newDocId = newDoc.getId();
    
    var doc = DocumentApp.openById(newDocId);
    var body = doc.getBody();

    var placeholders = {
      '{{Unique ID}}': clienteDados[2],
      '{{Nome Completo}}': clienteDados[3],
      '{{Estado Civil}}': clienteDados[4],
      '{{Profissão/Ocupação}}': clienteDados[5],
      '{{CPF ou CNPJ}}': clienteDados[6],
      '{{Número de RG}}': clienteDados[7],
      '{{Endereço Completo}}': clienteDados[8],
      '{{CEP da residência}}': clienteDados[9],
      '{{Número de Telefone}}': clienteDados[10],
      '{{Email}}': clienteDados[11],
      '{{valor_honocontrato}}': valor_honocontrato,
      '{{valor_honocontrato_extenso}}': valor_honocontrato_extenso,
      '{{n_parcelas}}': n_parcelas,
      '{{valorparcelado_honocontrato}}': valorparcelado_honocontrato,
      '{{valorparcelado_honocontrato_extenso}}': valorparcelado_honocontrato_extenso,
      '{{texto_pagamento}}': texto_pagamento,
    };

    for (var key in placeholders) {
      body.replaceText(key, placeholders[key]);
    }

    var data = new Date();
    var dia = Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd');
    var mes = obterMesEmPortugues(data.getMonth() + 1);
    var ano = Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy');
    var dataCriacao = dia + ' de ' + mes + ' de ' + ano;
    body.replaceText('{{Data}}', dataCriacao);

    doc.saveAndClose();

    folder.addFile(DriveApp.getFileById(newDocId));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(newDocId));

    Logger.log("Documento preenchido e salvo para o cliente " + clienteDados[3] + ": " + newDocName);
  }

  function obterOuCriarPasta(clienteDados) {
    var folderName = clienteDados[2] + ' - ' + clienteDados[3];
    var folders = DriveApp.getFolderById(rootFolderId).getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      var newFolder = DriveApp.getFolderById(rootFolderId).createFolder(folderName);
      return newFolder;
    }
  }

  var clienteFolder = obterOuCriarPasta(clienteDados);

  preencherEDatarDocumento(docTemplateId1, clienteDados, 'Procuracao', clienteFolder);
  preencherEDatarDocumento(docTemplateId2, clienteDados, 'CoHo', clienteFolder);
}

function formatarValor(valor) {
  return Utilities.formatString('R$ %s', valor.toFixed(2).replace('.', ','));
}

function converterParaExtenso(valor) {
  var unidades = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove'];
  var dezenas = ['dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
  var especiais = ['onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];
  var centenas = ['cem', 'cento e', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos'];

  function converteInteiro(numero) {
    var texto = '';

    if (numero === 100) return 'cem';
    
    if (numero >= 1000) {
      var milhar = Math.floor(numero / 1000);
      texto += (milhar > 1 ? unidades[milhar] + ' mil ' : 'mil ');
      numero %= 1000;
    }

    if (numero >= 100) {
      var centena = Math.floor(numero / 100);
      texto += (centena === 1 && numero % 100 === 0 ? 'cem ' : centenas[centena] + (numero % 100 > 0 ? ' e ' : ' '));
      numero %= 100;
    } else if (numero > 0 && texto !== '') {
      texto += 'e ';
    }

    if (numero >= 20) {
      var dezena = Math.floor(numero / 10);
      texto += dezenas[dezena - 1];
      numero %= 10;
      if (numero > 0) texto += ' e ' + unidades[numero];
    } else if (numero >= 11) {
      texto += especiais[numero - 11];
      numero = 0;
    } else if (numero > 0) {
      texto += unidades[numero];
    }

    return texto.trim();
  }

  var inteiro = Math.floor(valor);
  var decimal = Math.round((valor - inteiro) * 100);

  var extenso = converteInteiro(inteiro);
  if (inteiro !== 0) extenso += ' reais';
  
  if (decimal > 0) {
    var centavosTexto = converterCentavos(decimal);
    extenso += ' e ' + centavosTexto + ' centavos';
  }

  return extenso.trim();
}

function converterCentavos(decimal) {
  var unidades = ['zero', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove'];
  var dezenas = ['dez', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'];
  var especiais = ['onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'];

  if (decimal < 10) {
    return unidades[decimal];
  } else if (decimal < 20) {
    return especiais[decimal - 11];
  } else {
    var dezena = Math.floor(decimal / 10);
    var unidade = decimal % 10;
    var texto = dezenas[dezena - 1];
    if (unidade > 0) texto += ' e ' + unidades[unidade];
    return texto;
  }
}
