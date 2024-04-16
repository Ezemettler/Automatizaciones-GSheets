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
  var preciosFinales = pedidosSheet.getRange(2, 11, lastRow - 1, 1).getValues();  // 11: Columna precio final del producto
  var valoresEnvio = pedidosSheet.getRange(2, 21, lastRow - 1, 1).getValues();    // 21: Columna valor envio
  var valoresEmbalaje = pedidosSheet.getRange(2, 23, lastRow - 1, 1).getValues(); // 23: Columna valor embalaje
  var totalACobrar = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var total = preciosFinales[i][0] + valoresEnvio[i][0] + valoresEmbalaje[i][0];
    totalACobrar.push([total]);
  }
  pedidosSheet.getRange(2, 29, totalACobrar.length, 1).setValues(totalACobrar);   // 29: Columna total a cobrar
  
  // Actualizar el total cobrado
  var valoresSenado = pedidosSheet.getRange(2, 12, lastRow - 1, 1).getValues();   // 12: Columna valor seña
  var pagosRecibidos = pedidosSheet.getRange(2, 30, lastRow - 1, 1).getValues();  // 30: Columna pagos recibidos
  var totalCobrado = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var total = valoresSenado[i][0] + pagosRecibidos[i][0];
    totalCobrado.push([total]);
  }
  pedidosSheet.getRange(2, 31, totalCobrado.length, 1).setValues(totalCobrado);   // 31: Columna total cobrado
  
  // Actualizar el saldo a cobrar
  var totalACobrarValues = pedidosSheet.getRange(2, 29, lastRow - 1, 1).getValues();  // 29: Columna total a cobrar
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
    var porcentaje = (saldo / total) * 100;
    var etiqueta = porcentaje < 50 ? "Amarillo" : "Blanco";
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
        var fechaEntrega = formatDate(row[21]);   // Formatea la fecha de entrega de la columna 22 (index 21)
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
    var fechasOrdenadas = fechasAmarillo.concat(fechasBlanco).filter((fecha, index, self) => self.indexOf(fecha) === index).sort();
    
    // Define los encabezados de la hoja "Calendario Pedidos"
    var headers = ["Fecha"];    // Encabezado 1er columna
    for (var i = 0; i < fechasOrdenadas.length; i++) {    // Agrega encabezados para cada pedido según el número de fechas ordenadas
        headers.push("Pedido " + (i + 1));
    }

    calendarioSheet.getRange(1, 1, 1, headers.length).setValues([headers]);   // Establece los encabezados en la primera fila de la hoja "Calendario Pedidos"
    pintarCeldas(calendarioSheet, datosPorFechaAmarillo, datosPorFechaBlanco, fechasOrdenadas);   // Llama a la función que pinta las celdas de la hoja según los datos de pedidos
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


function pintarCeldas(hoja, datosPorFechaAmarillo, datosPorFechaBlanco, fechasOrdenadas) {
    for (var k = 0; k < fechasOrdenadas.length; k++) {    // Recorre cada fecha ordenada en la lista de fechas ordenadas
        var fecha = fechasOrdenadas[k];   // Obtiene la fecha actual del índice k
        var currentColumn = 2;    // Inicializa la columna actual en la segunda columna (donde se inician los pedidos)

        // Procesa pedidos con etiqueta amarilla
        if (datosPorFechaAmarillo[fecha]) {
            for (var l = 0; l < datosPorFechaAmarillo[fecha].length; l++) {   // Recorre cada pedido en la fecha actual con etiqueta amarilla
                var pedido = datosPorFechaAmarillo[fecha][l];   // Obtiene el pedido actual de la lista de pedidos amarillos para la fecha actual
                var espacioEnCalendario = pedido.espacioEnCalendario || 1;   // Obtiene el espacio en calendario del pedido, por defecto 1 si no está especificado

                // Define el rango de celdas para el pedido actual
                var rangeToSet = hoja.getRange(k + 2, currentColumn, 1, espacioEnCalendario);

                // Establece el valor del pedido en el rango de celdas
                rangeToSet.setValue(pedido.pedidoData);

                // Combina celdas si el espacio en calendario es mayor a 1
                if (espacioEnCalendario > 1) {
                    rangeToSet.merge();
                }

                // Establece el color de fondo de la celda según el estado del pedido
                if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
                    rangeToSet.setBackground("#34a853"); // Verde para pedido terminado
                } else {
                    rangeToSet.setBackground("#ffff00"); // Amarillo para pedidos con etiqueta amarilla
                }

                // Avanza la columna actual según el espacio en calendario
                currentColumn += espacioEnCalendario;
            }
        }

        // Procesa pedidos sin etiqueta amarilla
        if (datosPorFechaBlanco[fecha]) {
            for (var m = 0; m < datosPorFechaBlanco[fecha].length; m++) {   // Recorre cada pedido en la fecha actual sin etiqueta amarilla
                var pedido = datosPorFechaBlanco[fecha][m];   // Obtiene el pedido actual de la lista de pedidos blancos para la fecha actual
                var espacioEnCalendario = pedido.espacioEnCalendario || 1;   // Obtiene el espacio en calendario del pedido, por defecto 1 si no está especificado

                // Define el rango de celdas para el pedido actual
                var rangeToSet = hoja.getRange(k + 2, currentColumn, 1, espacioEnCalendario);

                // Establece el valor del pedido en el rango de celdas
                rangeToSet.setValue(pedido.pedidoData);

                // Combina celdas si el espacio en calendario es mayor a 1
                if (espacioEnCalendario > 1) {
                    rangeToSet.merge();
                }

                // Establece el color de fondo de la celda según el estado del pedido
                if (pedido.terminado && pedido.terminado.toLowerCase() === "si") {
                    rangeToSet.setBackground("#34a853"); // Verde para pedido terminado
                } else {
                    rangeToSet.setBackground("#ffffff"); // Blanco para pedidos sin etiqueta amarilla
                }

                // Avanza la columna actual según el espacio en calendario
                currentColumn += espacioEnCalendario;
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
  var rango = hoja.getRange("A1:H" + lastRow); // Definir el rango desde la columna A hasta la H y todas las filas registradas
  rango.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID); // Establecer los bordes negros
}