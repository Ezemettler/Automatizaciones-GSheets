function actualizarCalendarioPedidos() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var pedidosSheet = spreadsheet.getSheetByName("Pedidos");
  var calendarioSheet = spreadsheet.getSheetByName("Calendario Pedidos");
  
  if (!calendarioSheet) {
    calendarioSheet = spreadsheet.insertSheet("Calendario Pedidos");
  } else {
    calendarioSheet.clear();
  }

  var data = pedidosSheet.getDataRange().getValues();
  var datosPorFechaAmarillo = {};
  var datosPorFechaBlanco = {};

  for (var j = 1; j < data.length; j++) {
    var row = data[j];
    var fechaEntrega = formatDate(row[21]);
    var pedidoData = obtenerDatosPedido(row);
    var terminado = row[33];

    if (row[32].toLowerCase().includes("amarillo")) {
      if (!datosPorFechaAmarillo[fechaEntrega]) {
        datosPorFechaAmarillo[fechaEntrega] = [];
      }
      datosPorFechaAmarillo[fechaEntrega].push({ pedidoData: pedidoData, terminado: terminado });
    } else {
      if (!datosPorFechaBlanco[fechaEntrega]) {
        datosPorFechaBlanco[fechaEntrega] = [];
      }
      datosPorFechaBlanco[fechaEntrega].push({ pedidoData: pedidoData, terminado: terminado });
    }
  }

  var fechasAmarillo = Object.keys(datosPorFechaAmarillo);
  var fechasBlanco = Object.keys(datosPorFechaBlanco);
  var fechasOrdenadas = fechasAmarillo.concat(fechasBlanco).filter((fecha, index, self) => self.indexOf(fecha) === index).sort();
  var headers = ["Fecha"];

  for (var i = 0; i < fechasOrdenadas.length; i++) {
    headers.push("Pedido " + (i + 1));
  }
  
  calendarioSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  pintarCeldas(calendarioSheet, datosPorFechaAmarillo, datosPorFechaBlanco, fechasOrdenadas);

  calendarioSheet.autoResizeColumns(1, calendarioSheet.getLastColumn());

  alinearCeldasVerticalmente();
  formatoCabecera();
  agregarBordesNegros();
}


function pintarCeldas(hoja, datosPorFechaAmarillo, datosPorFechaBlanco, fechasOrdenadas) {
  for (var k = 0; k < fechasOrdenadas.length; k++) {
    var fecha = fechasOrdenadas[k];
    var rowData = [fecha];

    if (datosPorFechaAmarillo[fecha]) {
      for (var l = 0; l < datosPorFechaAmarillo[fecha].length; l++) {
        var pedido = datosPorFechaAmarillo[fecha][l];
        rowData.push(pedido.pedidoData);
        
        if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
          hoja.getRange(k + 2, l + 2).setBackground("#34a853");
        } else {
          hoja.getRange(k + 2, l + 2).setBackground("#ffff00");
        }
      }
    }
    
    if (datosPorFechaBlanco[fecha]) {
      for (var m = 0; m < datosPorFechaBlanco[fecha].length; m++) {
        var pedido = datosPorFechaBlanco[fecha][m];
        rowData.push(pedido.pedidoData);
        
        if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
          hoja.getRange(k + 2, datosPorFechaBlanco[fecha].length + m + 2).setBackground("#34a853");
        } else {
          hoja.getRange(k + 2, datosPorFechaBlanco[fecha].length + m + 2).setBackground("#ffffff");
        }
      }
    }
    
    hoja.appendRow(rowData);
  }
}


function obtenerDatosPedido(row) {
  // Esta función obtiene los datos de un pedido a partir de una fila de datos
  var comprador = row[14]; // Columna Comprador
  var producto = row[3]; // Columna 4: Tipo de producto
  var medidas = row[4]; // Columna 5: Medidas del producto
  var tela = row[5]; // Columna 6: Tela del producto
  var color = row[6]; // Columna 7: Color del producto
  var placa = row[7]; // Columna 8: Placa del producto
  var patas = row[8]; // Columna 9: Patas del producto
  var accesorios = row[9]; // Columna 10: Accesorios del producto
  var etiqueta = row[32]; // Columna 32: Etiqueta de color

  // Formatear los datos del pedido en una cadena de texto
  var pedidoText = comprador + "\n" + producto + " - " + medidas + "\n" + tela + " - " + color + "\nPlaca: " + placa + "\nPatas: " + patas + "\n" + accesorios + "\nEtiqueta: " + etiqueta;
  return pedidoText;
}

function formatDate(date) {
  // Esta función formatea una fecha en el formato "dd/MM/yyyy"
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return formattedDate;
}

function alinearCeldasVerticalmente() {
  // Esta función alinea verticalmente las celdas en la hoja de "Calendario Pedidos"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();   
  var hoja = spreadsheet.getSheetByName("Calendario Pedidos"); // Nombre de la hoja de calendario de pedidos
  var rango = hoja.getDataRange();      // Obtener el rango de celdas que deseas alinear verticalmente
  rango.setVerticalAlignment("middle"); // Alinear vertical las celdas
}


function formatoCabecera() {
  // Esta función establece el formato de la cabecera en la hoja de "Calendario Pedidos"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Calendario Pedidos"); // Nombre de la hoja de calendario de pedidos
  var cabeceraRange = hoja.getRange("1:1"); // Rango que cubre la fila de cabecera

  // Modificar el formato del texto en la fila de cabecera
  cabeceraRange.setFontSize(12); // Cambiar el tamaño de la letra
  cabeceraRange.setFontWeight("bold"); // Poner el texto en negrita

  // Inmovilizar la fila de cabecera
  hoja.setFrozenRows(1);
}

function agregarBordesNegros() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Calendario Pedidos"); // Nombre de la hoja de calendario de pedidos
  var lastRow = hoja.getLastRow();
  var rango = hoja.getRange("A1:H" + lastRow); // Definir el rango desde la columna A hasta la H y todas las filas registradas
  rango.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID); // Establecer los bordes negros
}