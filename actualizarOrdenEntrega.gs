function actualizarOrdenEntrega() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ordenEntregaSheet = ss.getSheetByName("Orden Entrega");
  var pedidosSheet = ss.getSheetByName("Pedidos");

  var numeroPedido = ordenEntregaSheet.getRange("F5").getValue(); // Obtener el número de pedido ingresado
  var columnaNumeroPedido = 25; // Buscar el número de pedido en la columna de la hoja Pedidos
  var ultimaFila = pedidosSheet.getLastRow();
  var data = pedidosSheet.getRange(2, columnaNumeroPedido, ultimaFila - 1).getValues(); // Devuelve una lista de todos los nro_pedidos
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === numeroPedido) {
      filaEncontrada = i + 2; // Guarda la fila cuando encuentra el nro de pedido en esa fila.
      break;
    }
  }

  if (filaEncontrada) {
    var comprador = pedidosSheet.getRange(filaEncontrada, 15).getValue(); // Obtener el valor del campo
    ordenEntregaSheet.getRange("B7").setValue(comprador); // Actualizar la celda en la hoja "Orden Entrega" con el valor del campo
    ordenEntregaSheet.getRange("B14").setValue(comprador); 
    var forma_pago = pedidosSheet.getRange(filaEncontrada, 14).getValue(); 
    ordenEntregaSheet.getRange("F10").setValue(forma_pago); 
    var forma_entrega = pedidosSheet.getRange(filaEncontrada, 17).getValue(); 
    ordenEntregaSheet.getRange("B10").setValue(forma_entrega); 
    var telefono = pedidosSheet.getRange(filaEncontrada, 16).getValue(); 
    ordenEntregaSheet.getRange("B15").setValue(telefono); 
    var direccion = pedidosSheet.getRange(filaEncontrada, 18).getValue(); 
    ordenEntregaSheet.getRange("B16").setValue(direccion); 
    var localidad = pedidosSheet.getRange(filaEncontrada, 19).getValue(); 
    var provincia = pedidosSheet.getRange(filaEncontrada, 20).getValue(); 
    var locacionCompleta = localidad + ", " + provincia;
    ordenEntregaSheet.getRange("B17").setValue(locacionCompleta);
    var producto = pedidosSheet.getRange(filaEncontrada, 4).getValue(); 
    var medidas = pedidosSheet.getRange(filaEncontrada, 5).getValue(); 
    var producto_medida = producto + " " + medidas;
    ordenEntregaSheet.getRange("B21").setValue(producto_medida);
    var tela = pedidosSheet.getRange(filaEncontrada, 6).getValue(); 
    var color = pedidosSheet.getRange(filaEncontrada, 7).getValue(); 
    var tela_color = tela + " " + color;
    ordenEntregaSheet.getRange("B22").setValue(tela_color);
    var placas = pedidosSheet.getRange(filaEncontrada, 8).getValue(); 
    var patas = pedidosSheet.getRange(filaEncontrada, 9).getValue(); 
    var placas_patas = "Placas " + placas + " - Patas " + patas;
    ordenEntregaSheet.getRange("B23").setValue(placas_patas);
    var accesorios = pedidosSheet.getRange(filaEncontrada, 10).getValue(); 
    ordenEntregaSheet.getRange("B24").setValue(accesorios); 
    var precio_producto = pedidosSheet.getRange(filaEncontrada, 11).getValue(); 
    ordenEntregaSheet.getRange("G28").setValue(precio_producto); 
    var valor_envio = pedidosSheet.getRange(filaEncontrada, 21).getValue(); 
    ordenEntregaSheet.getRange("G29").setValue(valor_envio); 
    var valor_embalaje = pedidosSheet.getRange(filaEncontrada, 23).getValue(); 
    ordenEntregaSheet.getRange("G30").setValue(valor_embalaje); 
    var senia = pedidosSheet.getRange(filaEncontrada, 12).getValue(); 
    ordenEntregaSheet.getRange("G32").setValue(senia);
    var recibe = pedidosSheet.getRange(filaEncontrada, 24).getValue(); 
    ordenEntregaSheet.getRange("F14").setValue(recibe);    
  } else {
    // Si el número de pedido no se encuentra, limpiar las celdas
    ordenEntregaSheet.getRange("B7").setValue("");
    ordenEntregaSheet.getRange("B14").setValue("");
    ordenEntregaSheet.getRange("B10").setValue("");
    ordenEntregaSheet.getRange("B15").setValue("");
    ordenEntregaSheet.getRange("B16").setValue("");
    ordenEntregaSheet.getRange("B17").setValue("");
    ordenEntregaSheet.getRange("B21").setValue("");
    ordenEntregaSheet.getRange("B22").setValue("");
    ordenEntregaSheet.getRange("B23").setValue("");
    ordenEntregaSheet.getRange("B24").setValue("");
    ordenEntregaSheet.getRange("G28").setValue("");
    ordenEntregaSheet.getRange("G29").setValue("");
    ordenEntregaSheet.getRange("G30").setValue("");
    ordenEntregaSheet.getRange("G32").setValue("");
    ordenEntregaSheet.getRange("F14").setValue("");
    ordenEntregaSheet.getRange("F10").setValue("");
  } 
}

