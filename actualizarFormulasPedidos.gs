function actualizarFormulasPedidos() {
  var pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  var lastRow = pedidosSheet.getLastRow();
  
  // Actualizar el número de pedido
  var numeroPedidoColumna = 26;
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
  var preciosFinales = pedidosSheet.getRange(2, 11, lastRow - 1, 1).getValues();
  var valoresEnvio = pedidosSheet.getRange(2, 21, lastRow - 1, 1).getValues();
  var valoresEmbalaje = pedidosSheet.getRange(2, 23, lastRow - 1, 1).getValues();
  var totalACobrar = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var total = preciosFinales[i][0] + valoresEnvio[i][0] + valoresEmbalaje[i][0];
    totalACobrar.push([total]);
  }
  pedidosSheet.getRange(2, 28, totalACobrar.length, 1).setValues(totalACobrar);
  
  // Actualizar el total cobrado
  var valoresSenado = pedidosSheet.getRange(2, 12, lastRow - 1, 1).getValues();
  var pagosRecibidos = pedidosSheet.getRange(2, 29, lastRow - 1, 1).getValues();
  var totalCobrado = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var total = valoresSenado[i][0] + pagosRecibidos[i][0];
    totalCobrado.push([total]);
  }
  pedidosSheet.getRange(2, 30, totalCobrado.length, 1).setValues(totalCobrado);
  
  // Actualizar el saldo a cobrar
  var totalACobrarValues = pedidosSheet.getRange(2, 28, lastRow - 1, 1).getValues();
  var totalCobradoValues = pedidosSheet.getRange(2, 30, lastRow - 1, 1).getValues();
  var saldoACobrar = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var saldo = totalACobrarValues[i][0] - totalCobradoValues[i][0];
    saldoACobrar.push([saldo]);
  }
  pedidosSheet.getRange(2, 31, saldoACobrar.length, 1).setValues(saldoACobrar);
  
  // Actualizar la etiqueta
  var etiquetas = [];
  for (var i = 0; i < lastRow - 1; i++) {
    var saldo = saldoACobrar[i][0];
    var total = totalACobrarValues[i][0];
    var porcentaje = (saldo / total) * 100;
    var etiqueta = porcentaje < 50 ? "Amarillo" : "Blanco";
    etiquetas.push([etiqueta]);
  }
  pedidosSheet.getRange(2, 32, etiquetas.length, 1).setValues(etiquetas);
}
