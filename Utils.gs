function formatearFecha(fecha) {
  try {
    const f = new Date(fecha);
    if (isNaN(f.getTime())) {
      throw new Error("Fecha inválida");
    }
    return Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (error) {
    console.error("Error en formatearFecha:", error);
    return "Fecha inválida";
  }
}

function getTipoMovimientoTexto(tipo) {
  switch (tipo.toUpperCase()) {
    case TIPOS_MOVIMIENTO.INGRESO:
      return "Ingreso";
    case TIPOS_MOVIMIENTO.SALIDA:
      return "Salida";
    case TIPOS_MOVIMIENTO.AJUSTE_POSITIVO:
      return "Ajuste Positivo";
    case TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO:
      return "Ajuste Negativo";
    case TIPOS_MOVIMIENTO.AJUSTE:
      return "Ajuste";
    case TIPOS_MOVIMIENTO.VENTA:
      return "Venta";
    default:
      return tipo;
  }
}

function parsearItems(itemsTexto) {
  const items = [];
  const partes = itemsTexto.split(',');
  
  for (const parte of partes) {
    const [codigo, cantidad] = parte.trim().split(':');
    if (codigo && cantidad) {
      items.push({
        codigo: codigo.trim(),
        cantidad: parseFloat(cantidad) || 0
      });
    }
  }
  
  return items;
}

function obtenerNombreProducto(codigo, datosProductos) {
  for (let i = 1; i < datosProductos.length; i++) {
    if (datosProductos[i][0] === codigo) {
      return datosProductos[i][1] || codigo;
    }
  }
  return codigo;
}

function obtenerCostoProducto(codigo, datosProductos) {
  for (let i = 1; i < datosProductos.length; i++) {
    if (datosProductos[i][0] === codigo) {
      // Si no hay precio en Productos, buscar en Inventario
      return 0; // Por ahora retornamos 0, se puede mejorar
    }
  }
  return 0;
}

function calcularValorInventario(datosInventario) {
  let valor = 0;
  
  for (let i = 1; i < datosInventario.length; i++) {
    const cantidad = Number(datosInventario[i][2]) || 0;
    const precio = Number(datosInventario[i][5]) || 0;
    valor += cantidad * precio;
  }
  
  return valor;
}

function generarDatosVacíos() {
  return {
    kpis: {
      ventasTotales: 0,
      ventasMes: 0,
      productosUnicos: 0,
      stockTotal: 0,
      margenPromedio: 0,
      rotacionInventario: 0,
      ticketPromedio: 0,
      disponibilidad: 0
    },
    ventasMensuales: [],
    topProductos: [],
    stockPorUbicacion: [],
    alertasStock: [],
    mejoresVendedores: [],
    topLugares: [],
    recomendaciones: [{
      tipo: 'info',
      texto: 'No hay datos suficientes para generar el análisis'
    }]
  };
}