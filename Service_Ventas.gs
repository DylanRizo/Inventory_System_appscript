function registrarVentaDetallada(datosVenta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    
    // Si no existe la hoja de ventas, crearla
    if (!hojaVentas) {
      hojaVentas = ss.insertSheet(HOJA_VENTAS);
      const encabezados = [[
        "ID Venta",
        "Fecha",
        "Hora Salida",
        "Hora Finalización",
        "Vendedor",
        "Entregador",
        "Items Vendidos",
        "Monto Cobrado",
        "Envío Cobrado",
        "Total",
        "Lugar Extracción",
        "Lugar Entrega",
        "Observaciones",
        "Timestamp"
      ]];
      hojaVentas.getRange(1, 1, 1, 14).setValues(encabezados);
      hojaVentas.getRange(1, 1, 1, 14)
        .setBackground("#dc3545")
        .setFontColor("white")
        .setFontWeight("bold");
      hojaVentas.setFrozenRows(1);
      hojaVentas.autoResizeColumns(1, 14);
    }
    
    // ========================================
    // VALIDACIONES BÁSICAS
    // ========================================
    if (!datosVenta.items || datosVenta.items.length === 0) {
      return { success: false, message: "❌ Debe incluir al menos un producto en la venta." };
    }
    
    if (!datosVenta.vendedor || !datosVenta.montoCobrado) {
      return { success: false, message: "❌ Vendedor y monto cobrado son obligatorios." };
    }

    if (!datosVenta.lugarExtraccion) {
      return { success: false, message: "❌ Debe especificar el lugar de extracción del inventario." };
    }
    
    // ========================================
    // PASO 1: SIMULACIÓN - VALIDAR STOCK DE TODOS LOS PRODUCTOS
    // ========================================
    Logger.log("🔍 PASO 1: Validando stock de todos los productos...");
    
    const validacionesStock = [];
    
    for (const item of datosVenta.items) {
      const codigo = item.codigo.toString().trim().toUpperCase();
      const cantidad = parseFloat(item.cantidad);
      
      // Validar que la cantidad sea válida
      if (cantidad <= 0 || isNaN(cantidad)) {
        return { 
          success: false, 
          message: `❌ Cantidad inválida para el producto ${codigo}. Debe ser mayor a 0.` 
        };
      }
      
      // Verificar stock disponible en la ubicación especificada
      const stockDisponible = verificarStockEnUbicacion(codigo, datosVenta.lugarExtraccion);
      
      Logger.log(`   📦 ${codigo}: Solicitado=${cantidad}, Disponible=${stockDisponible}`);
      
      if (stockDisponible < cantidad) {
        return { 
          success: false, 
          message: `❌ Stock insuficiente de "${codigo}" en "${datosVenta.lugarExtraccion}".\n\n` +
                   `📊 Disponible: ${stockDisponible}\n` +
                   `📦 Solicitado: ${cantidad}\n` +
                   `⚠️ Faltante: ${cantidad - stockDisponible}\n\n` +
                   `La venta NO ha sido registrada.`
        };
      }
      
      validacionesStock.push({
        codigo: codigo,
        cantidad: cantidad,
        stockDisponible: stockDisponible
      });
    }
    
    Logger.log("✅ PASO 1 COMPLETADO: Todos los productos tienen stock suficiente");
    
    // ========================================
    // PASO 2: EJECUCIÓN - PROCEDER CON LA TRANSACCIÓN
    // ========================================
    Logger.log("💾 PASO 2: Ejecutando transacción de venta...");
    
    // Generar ID único para la venta
    const timestamp = new Date();
    const idVenta = `V-${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss")}`;
    
    // Procesar items vendidos para el registro textual
    const itemsTexto = datosVenta.items.map(item => 
      `${item.codigo}:${item.cantidad}`
    ).join(", ");
    
    const montoCobrado = parseFloat(datosVenta.montoCobrado) || 0;
    const envioCobrado = parseFloat(datosVenta.envioCobrado) || 0;
    const total = montoCobrado + envioCobrado;
    
    // 2.1 - Registrar en hoja de Ventas
    Logger.log("   📝 Registrando venta en hoja Ventas...");
    hojaVentas.appendRow([
      idVenta,
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      datosVenta.horaSalida || "",
      datosVenta.horaFinalizacion || "",
      datosVenta.vendedor || "",
      datosVenta.entregador || "",
      itemsTexto,
      montoCobrado,
      envioCobrado,
      total,
      datosVenta.lugarExtraccion || "",
      datosVenta.lugarEntrega || "",
      datosVenta.observaciones || "",
      timestamp
    ]);
    
    Logger.log(`   ✅ Venta ${idVenta} registrada en hoja Ventas`);
    
    // 2.2 - Descontar inventario y registrar movimientos
    Logger.log("   📦 Descontando inventario y registrando movimientos...");
    const resultadosMovimientos = [];
    
    for (const validacion of validacionesStock) {
      const codigo = validacion.codigo;
      const cantidad = validacion.cantidad;
      
      // Descontar del inventario
      const resultadoDescuento = descontarDeInventario(
        codigo, 
        cantidad, 
        datosVenta.lugarExtraccion
      );
      
      if (!resultadoDescuento.success) {
        Logger.log(`   ⚠️ Error crítico al descontar ${codigo}: ${resultadoDescuento.message}`);
        return { 
          success: false, 
          message: `❌ Error crítico: La venta fue registrada pero falló el descuento de inventario para ${codigo}.\n` +
                   `Contacte al administrador. ID Venta: ${idVenta}` 
        };
      }
      
      Logger.log(`   ✅ Descontado ${cantidad} unidades de ${codigo}`);
      
      // Registrar movimiento
      const resultadoMov = registrarMovimiento({
        codigo: codigo,
        fecha: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tipo: TIPOS_MOVIMIENTO.VENTA,
        cantidad: cantidad,
        ubicacion: datosVenta.lugarExtraccion,
        observaciones: `Venta ${idVenta} - Vendedor: ${datosVenta.vendedor} - Entrega: ${datosVenta.lugarEntrega}`
      });
      
      if (!resultadoMov.includes("correctamente")) {
        Logger.log(`   ⚠️ Advertencia: El movimiento de ${codigo} no se registró correctamente`);
      }
      
      resultadosMovimientos.push(`✅ ${codigo}: ${cantidad} unidades`);
    }
    
    Logger.log("✅ PASO 2 COMPLETADO: Transacción ejecutada exitosamente");
    
    // ========================================
    // RESPUESTA EXITOSA
    // ========================================
    return {
      success: true,
      message: `✅ Venta registrada exitosamente.\n\n` +
               `📋 ID: ${idVenta}\n` +
               `👤 Vendedor: ${datosVenta.vendedor}\n` +
               `💰 Total: $${total.toFixed(2)}\n` +
               `📍 Extracción: ${datosVenta.lugarExtraccion}\n` +
               `📦 Entrega: ${datosVenta.lugarEntrega}\n\n` +
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
// FUNCIONES AUXILIARES (mantener sin cambios)
// ========================================

function obtenerReporteVentas(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    
    if (!hojaVentas || hojaVentas.getLastRow() <= 1) {
      return { ventas: [], kpis: {} };
    }
    
    const datos = hojaVentas.getDataRange().getValues();
    const ventas = [];
    
    // Convertir datos a objetos
    for (let i = 1; i < datos.length; i++) {
      const venta = {
        id: datos[i][0],
        fecha: datos[i][1],
        horaSalida: datos[i][2],
        horaFinalizacion: datos[i][3],
        vendedor: datos[i][4],
        entregador: datos[i][5],
        items: datos[i][6],
        montoCobrado: datos[i][7],
        envioCobrado: datos[i][8],
        total: datos[i][9],
        lugarExtraccion: datos[i][10],
        lugarEntrega: datos[i][11],
        observaciones: datos[i][12]
      };
      
      // Aplicar filtros si existen
      if (filtros) {
        if (filtros.fechaDesde && filtros.fechaHasta) {
          const fechaVenta = new Date(venta.fecha);
          const fechaDesde = new Date(filtros.fechaDesde);
          const fechaHasta = new Date(filtros.fechaHasta);
          
          if (fechaVenta < fechaDesde || fechaVenta > fechaHasta) {
            continue;
          }
        }
        
        if (filtros.vendedor && venta.vendedor !== filtros.vendedor) {
          continue;
        }
      }
      
      ventas.push(venta);
    }
    
    // Calcular KPIs
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
    ventasPorLugar
  };
}

function obtenerInfoVentaPorObservacion(observaciones, ventasSheet) {
  try {
    if (!observaciones || !ventasSheet) return null;
    
    // Buscar el ID de venta en las observaciones (formato: "Venta V-YYYYMMDD-HHMMSS")
    const match = observaciones.match(/Venta (V-\d{8}-\d{6})/);
    if (!match) return null;
    
    const idVenta = match[1];
    const datosVentas = ventasSheet.getDataRange().getValues();
    
    // Buscar la venta por ID
    for (let i = 1; i < datosVentas.length; i++) {
      if (datosVentas[i][0] === idVenta) {
        return {
          vendedor: datosVentas[i][4] || "N/A",
          entregador: datosVentas[i][5] || "N/A",
          lugarEntrega: datosVentas[i][11] || "N/A",
          montoTotal: datosVentas[i][9] || 0
        };
      }
    }
    
    return null;
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