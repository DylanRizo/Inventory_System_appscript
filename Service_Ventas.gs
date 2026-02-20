function registrarVentaDetallada(datosVenta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    
    // Si no existe la hoja de ventas, crearla con la nueva estructura
    if (!hojaVentas) {
      hojaVentas = ss.insertSheet(HOJA_VENTAS);
      const encabezados = [[
        "ID Venta",
        "Fecha",
        "Hora Salida",
        "Hora Finalización",
        "Vendedor",
        "Entregador",
        "Item",
        "Monto Cobrado",
        "Envío Cobrado",
        "Total",
        "Lugar Extracción",
        "Lugar Entrega",
        "Observaciones",
        "Timestamp",
        "Canal de Venta",
        "Precio Unitario"
      ]];
      hojaVentas.getRange(1, 1, 1, 16).setValues(encabezados);
      hojaVentas.getRange(1, 1, 1, 16)
        .setBackground("#dc3545")
        .setFontColor("white")
        .setFontWeight("bold");
      hojaVentas.setFrozenRows(1);
      hojaVentas.autoResizeColumns(1, 16);
    }
    
    // ========================================
    // VALIDACIONES BÁSICAS
    // ========================================
    if (!datosVenta.items || datosVenta.items.length === 0) {
      return { success: false, message: "❌ Debe incluir al menos un producto en la venta." };
    }
    
    if (!datosVenta.vendedor) {
      return { success: false, message: "❌ El vendedor es obligatorio." };
    }

    for (const item of datosVenta.items) {
      if (!item.almacen) {
        return { success: false, message: `❌ El producto "${item.codigo}" no tiene almacén de origen especificado.` };
      }
    }
    
    // ========================================
    // PASO 1: VALIDAR STOCK DE TODOS LOS PRODUCTOS
    // ========================================
    Logger.log("🔍 PASO 1: Validando stock de todos los productos...");
    
    const validacionesStock = [];
    
    for (const item of datosVenta.items) {
      const codigo = item.codigo.toString().trim().toUpperCase();
      const cantidad = parseFloat(item.cantidad);
      const almacen = item.almacen;
      
      if (cantidad <= 0 || isNaN(cantidad)) {
        return { 
          success: false, 
          message: `❌ Cantidad inválida para el producto ${codigo}. Debe ser mayor a 0.` 
        };
      }
      
      const stockDisponible = verificarStockEnUbicacion(codigo, almacen);
      
      Logger.log(`   📦 ${codigo} en ${almacen}: Solicitado=${cantidad}, Disponible=${stockDisponible}`);
      
      if (stockDisponible < cantidad) {
        return { 
          success: false, 
          message: `❌ Stock insuficiente de "${codigo}" en "${almacen}".\n\n` +
                   `📊 Disponible: ${stockDisponible}\n` +
                   `📦 Solicitado: ${cantidad}\n` +
                   `⚠️ Faltante: ${cantidad - stockDisponible}\n\n` +
                   `La venta NO ha sido registrada.`
        };
      }
      
      validacionesStock.push({
        codigo: codigo,
        nombre: item.nombre || codigo,
        cantidad: cantidad,
        precioUnit: parseFloat(item.precioUnit) || 0,
        subtotal: parseFloat(item.subtotal) || 0,
        almacen: almacen,
        stockDisponible: stockDisponible
      });
    }
    
    Logger.log("✅ PASO 1 COMPLETADO: Todos los productos tienen stock suficiente");
    
    // ========================================
    // PASO 2: EJECUTAR TRANSACCIÓN
    // ========================================
    Logger.log("💾 PASO 2: Ejecutando transacción de venta...");
    
    const timestamp = new Date();
    const idVenta = `V-${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss")}`;
    const fechaFormateada = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const envioCobrado = parseFloat(datosVenta.envioCobrado) || 0;
    const envioPorLinea = validacionesStock.length > 0
      ? Math.round((envioCobrado / validacionesStock.length) * 100) / 100
      : 0;
    
    // 2.1 - Registrar cada ítem como una fila en la hoja Ventas
    Logger.log("   📝 Registrando ítems en hoja Ventas...");
    Logger.log(`   🚚 Envío total: $${envioCobrado} / ${validacionesStock.length} ítems = $${envioPorLinea} por línea`);
    
    const filasVenta = [];
    
    for (const validacion of validacionesStock) {
      const totalLinea = validacion.subtotal + envioPorLinea;
      filasVenta.push([
        idVenta,                              // Col A - ID Venta
        fechaFormateada,                       // Col B - Fecha
        datosVenta.horaSalida || "",           // Col C - Hora Salida
        datosVenta.horaFinalizacion || "",     // Col D - Hora Finalización
        datosVenta.vendedor || "",             // Col E - Vendedor
        datosVenta.entregador || "",           // Col F - Entregador
        validacion.codigo + ":" + validacion.cantidad, // Col G - Item (CODIGO:CANTIDAD)
        validacion.subtotal,                   // Col H - Monto Cobrado (precio * cantidad)
        envioPorLinea,                         // Col I - Envío Cobrado (prorrateado)
        totalLinea,                            // Col J - Total (subtotal + envío prorrateado)
        validacion.almacen,                    // Col K - Lugar Extracción (almacén origen)
        datosVenta.lugarEntrega || "",         // Col L - Lugar Entrega
        datosVenta.observaciones || "",        // Col M - Observaciones
        timestamp,                             // Col N - Timestamp
        datosVenta.canalVenta || "",           // Col O - Canal de Venta
        validacion.precioUnit                  // Col P - Precio Unitario
      ]);
    }
    
    // Escribir todas las filas de una sola vez (batch)
    if (filasVenta.length > 0) {
      const ultimaFila = hojaVentas.getLastRow();
      hojaVentas.getRange(ultimaFila + 1, 1, filasVenta.length, 16).setValues(filasVenta);
    }
    
    Logger.log(`   ✅ ${filasVenta.length} fila(s) registradas para venta ${idVenta}`);
    
    // 2.2 - Descontar inventario y registrar movimientos
    Logger.log("   📦 Descontando inventario y registrando movimientos...");
    const resultadosMovimientos = [];
    
    for (const validacion of validacionesStock) {
      const resultadoDescuento = descontarDeInventario(
        validacion.codigo, 
        validacion.cantidad, 
        validacion.almacen
      );
      
      if (!resultadoDescuento.success) {
        Logger.log(`   ⚠️ Error crítico al descontar ${validacion.codigo}: ${resultadoDescuento.message}`);
        return { 
          success: false, 
          message: `❌ Error crítico: La venta fue registrada pero falló el descuento de inventario para ${validacion.codigo} en ${validacion.almacen}.\n` +
                   `Contacte al administrador. ID Venta: ${idVenta}` 
        };
      }
      
      Logger.log(`   ✅ Descontado ${validacion.cantidad} de ${validacion.codigo} en ${validacion.almacen}`);
      
      const resultadoMov = registrarMovimiento({
        codigo: validacion.codigo,
        fecha: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tipo: TIPOS_MOVIMIENTO.VENTA,
        cantidad: validacion.cantidad,
        ubicacion: validacion.almacen,
        observaciones: `Venta ${idVenta} - Vendedor: ${datosVenta.vendedor} - Entrega: ${datosVenta.lugarEntrega || 'N/A'}`
      });
      
      if (!resultadoMov.includes("correctamente")) {
        Logger.log(`   ⚠️ Advertencia: El movimiento de ${validacion.codigo} no se registró correctamente`);
      }
      
      resultadosMovimientos.push(`✅ ${validacion.codigo}: ${validacion.cantidad} unid. x $${validacion.precioUnit.toFixed(2)} = $${validacion.subtotal.toFixed(2)} (${validacion.almacen})`);
    }
    
    Logger.log("✅ PASO 2 COMPLETADO: Transacción ejecutada exitosamente");
    
    // ========================================
    // RESPUESTA EXITOSA
    // ========================================
    const subtotalProductos = validacionesStock.reduce(function(sum, v) { return sum + v.subtotal; }, 0);
    const totalFinal = subtotalProductos + envioCobrado;
    
    return {
      success: true,
      message: `✅ Venta registrada exitosamente.\n\n` +
               `📋 ID: ${idVenta}\n` +
               `👤 Vendedor: ${datosVenta.vendedor}\n` +
               `💰 Subtotal: $${subtotalProductos.toFixed(2)}\n` +
               `🚚 Envío: $${envioCobrado.toFixed(2)}\n` +
               `💵 Total: $${totalFinal.toFixed(2)}\n\n` +
               `Productos:\n${resultadosMovimientos.join('\n')}`,
      idVenta: idVenta
    };
    
  } catch (error) {
    console.error("❌ Error en registrarVentaDetallada:", error);
    Logger.log("Error stack: " + error.stack);
    return { 
      success: false, 
      message: `❌ Error del sistema: ${error.message}\n\nLa venta NO ha sido registrada.` 
    };
  }
}

// ========================================
// FUNCIONES AUXILIARES
// ========================================

/**
 * Obtiene el reporte de ventas agrupando filas por ID_Venta.
 * 1 fila por ítem, envío prorrateado por línea.
 * Columnas: [0]ID, [1]Fecha, [2]HoraSalida, [3]HoraFin, [4]Vendedor, [5]Entregador,
 *           [6]Item, [7]MontoCobrado, [8]EnvíoPorLínea, [9]Total, [10]LugarExt,
 *           [11]LugarEntrega, [12]Obs, [13]Timestamp, [14]Canal, [15]PrecioUnit
 */
function obtenerReporteVentas(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    
    if (!hojaVentas || hojaVentas.getLastRow() <= 1) {
      return { ventas: [], kpis: {} };
    }
    
    const datos = hojaVentas.getDataRange().getValues();
    
    // Agrupar filas por ID_Venta para reconstruir ventas completas
    const ventasMap = {};
    
    for (let i = 1; i < datos.length; i++) {
      const idVenta = datos[i][0];
      if (!idVenta) continue;
      
      // Aplicar filtro de fechas
      if (filtros && filtros.fechaDesde && filtros.fechaHasta) {
        const fechaVenta = new Date(datos[i][1]);
        const fechaDesde = new Date(filtros.fechaDesde);
        const fechaHasta = new Date(filtros.fechaHasta);
        
        if (fechaVenta < fechaDesde || fechaVenta > fechaHasta) {
          continue;
        }
      }
      
      // Aplicar filtro de vendedor
      if (filtros && filtros.vendedor && datos[i][4] !== filtros.vendedor) {
        continue;
      }
      
      if (!ventasMap[idVenta]) {
        ventasMap[idVenta] = {
          id: idVenta,
          fecha: datos[i][1],
          horaSalida: datos[i][2],
          horaFinalizacion: datos[i][3],
          vendedor: datos[i][4],
          entregador: datos[i][5],
          lugarEntrega: datos[i][11],
          observaciones: datos[i][12],
          canalVenta: datos[i][14] || "No especificado",
          envioCobrado: 0,
          items: [],
          total: 0
        };
      }
      
      const montoCobrado = parseFloat(datos[i][7]) || 0;  // Col H - subtotal del ítem
      const envioPorLinea = parseFloat(datos[i][8]) || 0;  // Col I - envío prorrateado
      const totalLinea = parseFloat(datos[i][9]) || 0;     // Col J - total de la línea
      
      ventasMap[idVenta].envioCobrado += envioPorLinea;     // Sumar envío prorrateado
      ventasMap[idVenta].items.push({
        producto: datos[i][6],
        montoCobrado: montoCobrado,
        precioUnit: datos[i][15],
        subtotal: montoCobrado,
        envioPorLinea: envioPorLinea,
        totalLinea: totalLinea,
        almacen: datos[i][10]
      });
      ventasMap[idVenta].total += totalLinea;
    }
    
    // Convertir mapa a array
    const ventas = Object.values(ventasMap);
    
    const kpis = calcularKPIsVentas(ventas);
    
    return { ventas, kpis };
  } catch (error) {
    console.error("Error en obtenerReporteVentas:", error);
    return { ventas: [], kpis: {} };
  }
}

function calcularKPIsVentas(ventas) {
  if (!ventas || ventas.length === 0) {
    return {
      totalVentas: 0,
      montoTotal: 0,
      promedioVenta: 0,
      mejorVendedor: null,
      lugarMasVentas: null
    };
  }
  
  const ventasPorVendedor = {};
  const ventasPorLugar = {};
  const ventasPorCanal = {};
  let montoTotal = 0;
  
  ventas.forEach(venta => {
    // Ventas por vendedor
    if (!ventasPorVendedor[venta.vendedor]) {
      ventasPorVendedor[venta.vendedor] = { cantidad: 0, monto: 0 };
    }
    ventasPorVendedor[venta.vendedor].cantidad++;
    ventasPorVendedor[venta.vendedor].monto += venta.total;
    
    // Ventas por lugar
    if (!ventasPorLugar[venta.lugarEntrega]) {
      ventasPorLugar[venta.lugarEntrega] = { cantidad: 0, monto: 0 };
    }
    ventasPorLugar[venta.lugarEntrega].cantidad++;
    ventasPorLugar[venta.lugarEntrega].monto += venta.total;
    
    // Ventas por canal
    const canal = venta.canalVenta || "No especificado";
    if (!ventasPorCanal[canal]) {
      ventasPorCanal[canal] = { cantidad: 0, monto: 0 };
    }
    ventasPorCanal[canal].cantidad++;
    ventasPorCanal[canal].monto += venta.total;
    
    montoTotal += venta.total;
  });
  
  // Encontrar mejor vendedor
  let mejorVendedor = null;
  let maxMonto = 0;
  
  for (const vendedor in ventasPorVendedor) {
    if (ventasPorVendedor[vendedor].monto > maxMonto) {
      maxMonto = ventasPorVendedor[vendedor].monto;
      mejorVendedor = {
        nombre: vendedor,
        ventas: ventasPorVendedor[vendedor].cantidad,
        monto: ventasPorVendedor[vendedor].monto
      };
    }
  }
  
  // Encontrar lugar con más ventas
  let lugarMasVentas = null;
  let maxCantidad = 0;
  
  for (const lugar in ventasPorLugar) {
    if (ventasPorLugar[lugar].cantidad > maxCantidad) {
      maxCantidad = ventasPorLugar[lugar].cantidad;
      lugarMasVentas = {
        nombre: lugar,
        ventas: ventasPorLugar[lugar].cantidad,
        monto: ventasPorLugar[lugar].monto
      };
    }
  }
  
  return {
    totalVentas: ventas.length,
    montoTotal: Math.round(montoTotal * 100) / 100,
    promedioVenta: Math.round((montoTotal / ventas.length) * 100) / 100,
    mejorVendedor,
    lugarMasVentas,
    ventasPorVendedor,
    ventasPorLugar,
    ventasPorCanal
  };
}

/**
 * Busca info de venta a partir de las observaciones de un movimiento.
 * Busca por ID_Venta y opcionalmente por código de producto para obtener
 * el total correcto de cada ítem individual.
 * Col G = "CODE - nombre", Col J = Total de la línea
 */
function obtenerInfoVentaPorObservacion(observaciones, ventasSheet, productoCodigo) {
  try {
    if (!observaciones || !ventasSheet) return null;
    
    const match = observaciones.match(/Venta (V-\d{8}-\d{6})/);
    if (!match) return null;
    
    const idVenta = match[1];
    const datosVentas = ventasSheet.getDataRange().getValues();
    const codigoUpper = productoCodigo ? productoCodigo.toString().toUpperCase() : null;
    
    let primeraCoincidencia = null;
    
    for (let i = 1; i < datosVentas.length; i++) {
      if (datosVentas[i][0] !== idVenta) continue;
      
      const itemCol = (datosVentas[i][6] || "").toString().toUpperCase();
      const itemCodigoEnFila = itemCol.split(':')[0].trim();
      
      // Si tenemos código de producto, buscar la fila exacta
      if (codigoUpper && itemCodigoEnFila === codigoUpper) {
        return {
          vendedor: datosVentas[i][4] || "N/A",
          entregador: datosVentas[i][5] || "N/A",
          lugarEntrega: datosVentas[i][11] || "N/A",
          montoTotal: parseFloat(datosVentas[i][9]) || 0  // Col J - Total de ESTA línea
        };
      }
      
      // Guardar primera coincidencia como fallback
      if (!primeraCoincidencia) {
        primeraCoincidencia = {
          vendedor: datosVentas[i][4] || "N/A",
          entregador: datosVentas[i][5] || "N/A",
          lugarEntrega: datosVentas[i][11] || "N/A",
          montoTotal: parseFloat(datosVentas[i][9]) || 0
        };
      }
    }
    
    return primeraCoincidencia;
  } catch (error) {
    console.error("Error en obtenerInfoVentaPorObservacion:", error);
    return null;
  }
}

function obtenerVendedores() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ventasSheet = ss.getSheetByName(HOJA_VENTAS);
    
    if (!ventasSheet || ventasSheet.getLastRow() <= 1) {
      return [];
    }
    
    const datos = ventasSheet.getDataRange().getValues();
    const vendedoresSet = new Set();
    
    // Columna 4 es "Vendedor"
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][4] && datos[i][4].toString().trim() !== "") {
        vendedoresSet.add(datos[i][4].toString().trim());
      }
    }
    
    return Array.from(vendedoresSet).sort();
  } catch (error) {
    console.error("Error en obtenerVendedores:", error);
    return [];
  }
}