function funcionesOrdenadasPedidos() {
  actualizarFormulasPedidos()     // 1er funcion a ejecutar.
  actualizarCalendarioPedidos()   // 2da funcion a ejecutar.
}


function actualizarFormulasPedidos() {
// Actualiza los campos con formulas en la hoja Pedidos.   
var pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
var lastRow = pedidosSheet.getLastRow();

// Actualizar el número de pedido
var numeroPedidoColumna = 27;
var numeroPedidoRange = pedidosSheet.getRange(2, numeroPedidoColumna, lastRow - 1, 1);
var numerosPedido = numeroPedidoRange.getValues();
var maxNumeroPedido = Math.max.apply(null, numerosPedido.flat());
var primerNumeroPedido = isNaN(maxNumeroPedido) ? 1000001 : maxNumeroPedido + 1;
var nuevoNumeroPedido = primerNumeroPedido;
for (var i = 0; i < numerosPedido.length; i++) {
  if (numerosPedido[i][0] === "" || isNaN(numerosPedido[i][0])) {
    pedidosSheet.getRange(i + 2, numeroPedidoColumna).setValue(nuevoNumeroPedido++);
  }
}

// Actualizar el total a cobrar
var preciosFinales = pedidosSheet.getRange(2, 22, lastRow - 1, 1).getValues();  // 22: Columna precio final del producto
var valoresEnvio = pedidosSheet.getRange(2, 11, lastRow - 1, 1).getValues();    // 11: Columna valor envio
var valoresEmbalaje = pedidosSheet.getRange(2, 12, lastRow - 1, 1).getValues(); // 12: Columna valor embalaje
var totalACobrar = [];
for (var i = 0; i < lastRow - 1; i++) {
  var total = preciosFinales[i][0] + valoresEnvio[i][0] + valoresEmbalaje[i][0];
  totalACobrar.push([total]);
}
pedidosSheet.getRange(2, 29, totalACobrar.length, 1).setValues(totalACobrar);   // 29: Columna total a cobrar

// Actualizar el total cobrado
var valoresSenado = pedidosSheet.getRange(2, 23, lastRow - 1, 1).getValues();   // 23: Columna valor seña
var pagosRecibidos = pedidosSheet.getRange(2, 30, lastRow - 1, 1).getValues();  // 30: Columna pagos recibidos
var totalCobrado = [];
for (var i = 0; i < lastRow - 1; i++) {
  var total = valoresSenado[i][0] + pagosRecibidos[i][0];
  totalCobrado.push([total]);
}
pedidosSheet.getRange(2, 31, totalCobrado.length, 1).setValues(totalCobrado);   // 31: Columna total cobrado

// Actualizar el saldo a cobrar
var totalACobrarValues = pedidosSheet.getRange(2, 29, lastRow - 1, 1).getValues();  // 29: Columna total a cobrar
var espacioCalendarioValues = pedidosSheet.getRange(2, 21, lastRow - 1, 1).getValues();  // 21: Columna Espacio en el calendario
var totalCobradoValues = pedidosSheet.getRange(2, 31, lastRow - 1, 1).getValues();  // 31: Columna total cobrado
var saldoACobrar = [];
for (var i = 0; i < lastRow - 1; i++) {
  var saldo = totalACobrarValues[i][0] - totalCobradoValues[i][0];
  saldoACobrar.push([saldo]);
}
pedidosSheet.getRange(2, 32, saldoACobrar.length, 1).setValues(saldoACobrar);   // 32: Columna saldo a cobrar

  // Actualizar la etiqueta
var etiquetas = [];
for (var i = 0; i < lastRow - 1; i++) {
  var saldo = saldoACobrar[i][0];
  var total = totalACobrarValues[i][0];
  var espacioCalendario = espacioCalendarioValues[i][0]; // Reemplaza con el número de columna adecuado
  var etiqueta;
  if (espacioCalendario == 0) {
    etiqueta = "Terciarizado";
  } else {
    var porcentaje = (saldo / total) * 100;
    etiqueta = porcentaje < 50 ? "Amarillo" : "Blanco";
  }
  etiquetas.push([etiqueta]);
}
pedidosSheet.getRange(2, 33, etiquetas.length, 1).setValues(etiquetas);   // 33: Columna etiqueta
}


function actualizarCalendarioPedidos() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();    // Obtiene la hoja de cálculo activa
  var pedidosSheet = spreadsheet.getSheetByName("Pedidos");   // Obtiene la hoja "Pedidos" de la hoja de cálculo
  var calendarioSheet = spreadsheet.getSheetByName("Calendario Pedidos");   // Obtiene la hoja "Calendario Pedidos" de la hoja de cálculo

  if (!calendarioSheet) {
      calendarioSheet = spreadsheet.insertSheet("Calendario Pedidos");      // Si la hoja "Calendario Pedidos" no existe, la crea
  } else {
      calendarioSheet.clear();    // Si la hoja ya existe, la limpia completamente
  }

  var data = pedidosSheet.getDataRange().getValues();   // Obtiene los datos de la hoja "Pedidos"
  var datosPorFechaAmarillo = {}; // Crea diccionarios para almacenar pedidos por fecha y etiqueta
  var datosPorFechaBlanco = {};

  for (var j = 1; j < data.length; j++) {       // Itera sobre cada fila de datos de la hoja "Pedidos", a partir de la segunda fila (excluye la cabecera)
      var row = data[j];                        // Obtiene la fila actual del ciclo.
      var fechaEntrega = formatDate(row[12]);   // Formatea la fecha de entrega de la columna 13 (index 12)
      var pedidoData = obtenerDatosPedido(row); // Obtiene los datos de un pedido de la fila actual
      var terminado = row[33];                  // Obtiene el estado de finalización del pedido de la columna 34 (index 33)

      // Verifica si el pedido tiene la etiqueta "amarillo"
      if (row[32].includes("Amarillo")) {
          if (!datosPorFechaAmarillo[fechaEntrega]) {     // Si la fecha no está en el diccionario, la crea como un array vacío
              datosPorFechaAmarillo[fechaEntrega] = []; 
          }
          datosPorFechaAmarillo[fechaEntrega].push({ pedidoData: pedidoData, terminado: terminado });   // Agrega el pedido al diccionario correspondiente a la fecha
      } else {
          if (!datosPorFechaBlanco[fechaEntrega]) {       // Si la fecha no está en el diccionario, la crea como un array vacío
              datosPorFechaBlanco[fechaEntrega] = [];
          }
          datosPorFechaBlanco[fechaEntrega].push({ pedidoData: pedidoData, terminado: terminado });     // Agrega el pedido al diccionario correspondiente a la fecha
      }
  }

  // Obtiene las fechas de los pedidos con etiqueta "amarillo" y "blanco"
  var fechasAmarillo = Object.keys(datosPorFechaAmarillo);
  var fechasBlanco = Object.keys(datosPorFechaBlanco);
  
  // Combina y ordena las fechas sin duplicados
  var fechasOrdenadas = fechasAmarillo.concat(fechasBlanco).filter((fecha, index, self) => self.indexOf(fecha) === index).sort(function(a, b) {
      return new Date(a.split("/").reverse().join("/")) - new Date(b.split("/").reverse().join("/"));
  });
  
  // Define los encabezados de la hoja "Calendario Pedidos"
  var headers = ["Fecha"];    // Encabezado 1er columna
  for (var i = 0; i < fechasOrdenadas.length; i++) {    // Agrega encabezados para cada pedido según el número de fechas ordenadas
      headers.push("Pedido " + (i + 1));
  }

  // Define un diccionario para almacenar pedidos terciarizados por fecha
  var datosPorFechaTerciarizado = {};
  // Itera sobre los datos para identificar y almacenar los pedidos terciarizados
  for (var j = 1; j < data.length; j++) {
      var row = data[j]; // Obtiene la fila actual del ciclo
      var fechaEntrega = formatDate(row[12]); // Formatea la fecha de entrega
      var pedidoData = obtenerDatosPedido(row); // Obtiene los datos del pedido
      var terminado = row[33]; // Obtiene el estado de finalización del pedido

      // Verifica si el pedido está terciarizado
      if (row[32].includes("Terciarizado")) {
          if (!datosPorFechaTerciarizado[fechaEntrega]) { // Si la fecha no está en el diccionario, la crea como un array vacío
              datosPorFechaTerciarizado[fechaEntrega] = [];
          }
          datosPorFechaTerciarizado[fechaEntrega].push({ pedidoData: pedidoData, terminado: terminado }); // Agrega el pedido al diccionario correspondiente a la fecha
      }
  }


  calendarioSheet.getRange(1, 1, 1, headers.length).setValues([headers]);   // Establece los encabezados en la primera fila de la hoja "Calendario Pedidos"
  pintarCeldas(calendarioSheet, datosPorFechaAmarillo, datosPorFechaBlanco, datosPorFechaTerciarizado, fechasOrdenadas);  // Llama a la función que pinta las celdas según los datos de pedidos
  calendarioSheet.autoResizeColumns(1, calendarioSheet.getLastColumn());    // Ajusta automáticamente el ancho de las columnas

  alinearCeldasVerticalmente();   // Alinea verticalmente las celdas en la hoja "Calendario Pedidos"
  formatoCabecera();      // Aplica formato a la cabecera de la hoja "Calendario Pedidos"
  agregarBordesNegros();  // Agrega bordes negros a las celdas de la hoja "Calendario Pedidos"
}


function formatDate(date) {
// Esta función formatea una fecha en el formato "dd/MM/yyyy"
var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
return formattedDate;
}

function obtenerDatosPedido(row) {
// Esta función obtiene los datos de un pedido a partir de una fila de datos
var comprador = row[3]; // Columna Comprador
var producto = row[13]; // Columna Producto
var cantidad = row[14]; // Columna Cantidad
var tela = row[15]; // Columna Tela del producto
var color = row[16]; // Columna Color del producto
var placa = row[17]; // Columna Placa del producto
var patas = row[18]; // Columna Patas del producto
var detalle_pedido = row[19]; // Columna detalle del pedido
var etiqueta = row[32]; // Columna Etiqueta de color

// Formatear los datos del pedido en una cadena de texto
var pedidoText = comprador + "\n" + producto + " (" + cantidad + " Un)\n" + tela + " - " + color + "\nPlaca: " + placa + "\nPatas: " + patas + "\n" + detalle_pedido + "\nEtiqueta: " + etiqueta;
return pedidoText;
}


function pintarCeldas(hoja, datosPorFechaAmarillo, datosPorFechaBlanco, datosPorFechaTerciarizado, fechasOrdenadas) {
  // Imprimir las fechas ordenadas en la primera columna
  for (var i = 0; i < fechasOrdenadas.length; i++) {
      var fecha = fechasOrdenadas[i];
      // Establecer la fecha en la primera columna (columna A) y la fila i + 2 (para evitar sobrescribir el encabezado)
      hoja.getRange(i + 2, 1).setValue(fecha);
  }

  // El resto de la función se mantiene igual para procesar los pedidos con etiquetas amarillas y blancas
  for (var k = 0; k < fechasOrdenadas.length; k++) {    // Recorre cada fecha ordenada en la lista de fechas ordenadas.
      var fecha = fechasOrdenadas[k];   // Obtiene la fecha actual del índice k.
      var currentColumn = 2;    // Inicializa la columna actual en la segunda columna (donde se inician los pedidos).

      // Procesa pedidos con etiqueta amarilla.
      if (datosPorFechaAmarillo[fecha]) {
          for (var l = 0; l < datosPorFechaAmarillo[fecha].length; l++) {   // Recorre cada pedido en la fecha actual con etiqueta amarilla.
              var pedido = datosPorFechaAmarillo[fecha][l];   // Obtiene el pedido actual de la lista de pedidos amarillos para la fecha actual.

              // Define el rango de celdas para el pedido actual.
              var rangeToSet = hoja.getRange(k + 2, currentColumn, 1, 1); // Rango de celdas de 1x1.

              // Establece el valor del pedido en el rango de celdas.
              rangeToSet.setValue(pedido.pedidoData);

              // Establece el color de fondo de la celda según el estado del pedido.
              if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
                  rangeToSet.setBackground("#34a853"); // Verde para pedido terminado.
              } else {
                  rangeToSet.setBackground("#ffff00"); // Amarillo para pedidos con etiqueta amarilla.
              }

              // Avanza la columna actual.
              currentColumn++;
          }
      }

      // Procesa pedidos sin etiqueta amarilla.
      if (datosPorFechaBlanco[fecha]) {
          for (var m = 0; m < datosPorFechaBlanco[fecha].length; m++) {   // Recorre cada pedido en la fecha actual sin etiqueta amarilla.
              var pedido = datosPorFechaBlanco[fecha][m];   // Obtiene el pedido actual de la lista de pedidos blancos para la fecha actual.

              // Define el rango de celdas para el pedido actual.
              var rangeToSet = hoja.getRange(k + 2, currentColumn, 1, 1); // Rango de celdas de 1x1.

              // Establece el valor del pedido en el rango de celdas.
              rangeToSet.setValue(pedido.pedidoData);

              // Establece el color de fondo de la celda según el estado del pedido.
              if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
                  rangeToSet.setBackground("#34a853"); // Verde para pedido terminado.
              } else {
                  rangeToSet.setBackground("#ffffff"); // Blanco para pedidos sin etiqueta amarilla.
              }

              // Avanza la columna actual.
              currentColumn++;
          }
      }

      // Procesa pedidos "Terciarizados".
      if (datosPorFechaTerciarizado[fecha]) {
          for (var n = 0; n < datosPorFechaTerciarizado[fecha].length; n++) {   // Recorre cada pedido "Terciarizado" en la fecha actual.
              var pedido = datosPorFechaTerciarizado[fecha][n];   // Obtiene el pedido actual de la lista de pedidos "Terciarizados" para la fecha actual.

              // Define el rango de celdas para el pedido actual.
              var rangeToSet = hoja.getRange(k + 2, currentColumn, 1, 1); // Rango de celdas de 1x1.

              // Establece el valor del pedido en el rango de celdas.
              rangeToSet.setValue(pedido.pedidoData);

              // Establece el color de fondo de la celda para los pedidos "Terciarizados".
              rangeToSet.setBackground("#8e7cc3"); // Violeta para pedidos "Terciarizados".

              // Avanza la columna actual.
              currentColumn++;
          }
      }
  }
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

hoja.setFrozenRows(1);  // Inmovilizar la fila de cabecera
}

function agregarBordesNegros() {
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var hoja = spreadsheet.getSheetByName("Calendario Pedidos"); // Nombre de la hoja de calendario de pedidos
var lastRow = hoja.getLastRow();

// Define los anchos de columna deseados (en píxeles)
var columnWidths = {
  1: 75,   // Anchura de la columna A (Fecha)
  2: 250,  // Ancho de columna N°
  3: 250,   
  4: 250,
  5: 250,
  6: 250,
  7: 250,
  8: 250,
  9: 250
  // Agrega más anchos de columna según sea necesario
};

// Define el rango de celdas donde deseas aplicar el ajuste de texto
var rango = hoja.getRange("A1:J" + lastRow);

// Aplica el ajuste de texto al rango de celdas
rango.setWrap(true);

// Establece los anchos de columna
for (var columna in columnWidths) {
  hoja.setColumnWidth(parseInt(columna), columnWidths[columna]);
}

// Establece los bordes negros alrededor del rango de celdas
rango.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID); 
}



