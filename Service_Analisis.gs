function obtenerHistorial(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const ventasSheet = ss.getSheetByName(HOJA_VENTAS);
    
    if (!movSheet || !prodSheet) {
      throw new Error("Las hojas del sistema no existen.");
    }
    
    const movimientos = movSheet.getDataRange().getValues();
    const productos = prodSheet.getDataRange().getValues();
    
    if (movimientos.length <= 1) {
      return [];
    }
    
    // Crear mapa de nombres de productos
    const prodMap = {};
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0]) {
        prodMap[productos[i][0].toString().toUpperCase()] = productos[i][1];
      }
    }
    
    // Validar fechas
    const fechaDesde = new Date(filtros.fechaDesde + 'T00:00:00');
    const fechaHasta = new Date(filtros.fechaHasta + 'T23:59:59');
    
    if (fechaDesde > fechaHasta) {
      throw new Error("La fecha 'desde' no puede ser posterior a la fecha 'hasta'");
    }
    
    const resultado = [];
    
    // Leer encabezados de movimientos
    const headersMov = movimientos[0];
    const colMov = {};
    headersMov.forEach((h, i) => {
      colMov[h.toString().trim()] = i;
    });
    
    Logger.log("📋 Encabezados Movimientos: " + JSON.stringify(colMov));
    
    // Procesar movimientos con filtros avanzados
    for (let i = 1; i < movimientos.length; i++) {
      const mov = movimientos[i];
      if (!mov[colMov['Código']] || !mov[colMov['Fecha']]) continue;
      
      try {
        const fechaMov = new Date(mov[colMov['Fecha']]);
        const tipoMov = mov[colMov['Tipo']] ? mov[colMov['Tipo']].toString().toUpperCase() : "";
        const codigoProducto = mov[colMov['Código']].toString().toUpperCase();
        const ubicacionCol = colMov['Ubicación'] !== undefined ? colMov['Ubicación'] : 8;
        const ubicacionMov = mov[ubicacionCol] ? mov[ubicacionCol].toString().trim() : "";
        
        // FILTRO POR FECHA
        if (fechaMov < fechaDesde || fechaMov > fechaHasta) {
          continue;
        }
        
        // FILTRO POR TIPO DE MOVIMIENTO
        if (filtros.tipo && tipoMov !== filtros.tipo.toUpperCase()) {
          continue;
        }
        
        // FILTRO POR UBICACIÓN
        if (filtros.ubicacion && filtros.ubicacion !== "" && ubicacionMov !== filtros.ubicacion) {
          continue;
        }
        
        // FILTRO POR PRODUCTO
        if (filtros.producto && filtros.producto !== "" && codigoProducto !== filtros.producto.toUpperCase()) {
          continue;
        }
        
        // Construir registro
        const registro = {
          codigo: mov[colMov['Código']],
          fecha: formatearFecha(fechaMov),
          tipo: tipoMov,
          cantidad: parseFloat(mov[colMov['Cantidad']]) || 0,
          producto: prodMap[codigoProducto] || "Producto no encontrado",
          observaciones: mov[colMov['Observaciones']] || "",
          usuario: mov[colMov['Usuario']] || "N/A",
          ubicacion: ubicacionMov || "-"
        };
        
        // Si es una VENTA, intentar obtener información del vendedor
        if (tipoMov === 'VENTA' && ventasSheet) {
          const infoVenta = obtenerInfoVentaPorObservacion(mov[colMov['Observaciones']], ventasSheet);
          if (infoVenta) {
            registro.vendedor = infoVenta.vendedor;
            registro.entregador = infoVenta.entregador;
            registro.lugarEntrega = infoVenta.lugarEntrega;
            registro.montoTotal = infoVenta.montoTotal;
            
            // FILTRO POR VENDEDOR (solo aplica a ventas)
            if (filtros.vendedor && filtros.vendedor !== "" && registro.vendedor !== filtros.vendedor) {
              continue;
            }
          }
        }
        
        resultado.push(registro);
        
      } catch (dateError) {
        console.warn(`Fecha inválida en movimiento fila ${i + 1}:`, mov[colMov['Fecha']]);
        continue;
      }
    }
    
    Logger.log(`✅ Historial generado: ${resultado.length} movimientos encontrados`);
    
    // Ordenar por fecha descendente
    return resultado.sort((a, b) => {
      const fechaA = new Date(a.fecha.split('/').reverse().join('-'));
      const fechaB = new Date(b.fecha.split('/').reverse().join('-'));
      return fechaB - fechaA;
    });
    
  } catch (error) {
    console.error("Error en obtenerHistorial:", error);
    Logger.log("Error stack: " + error.stack);
    return [];
  }
}

function obtenerResumen() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    if (!prodSheet || !movSheet) {
      return { totalProductos: 0, totalMovimientos: 0, sinStock: 0, stockBajo: 0, valorTotalInventario: 0 };
    }
    
    const productos = prodSheet.getDataRange().getValues();
    const movimientos = movSheet.getDataRange().getValues();
    
    const totalProductos = Math.max(0, productos.length - 1);
    const totalMovimientos = Math.max(0, movimientos.length - 1);
    
    let sinStock = 0;
    let stockBajo = 0;
    let valorTotalInventario = 0;
    
    for (let i = 1; i < productos.length; i++) {
      if (!productos[i][0]) continue;
      
      const codigo = productos[i][0];
      const stockMin = Math.max(0, parseInt(productos[i][4]) || 0);
      const stock = calcularStock(codigo);
      const precio = parseFloat(productos[i][5]) || 0;
      
      if (stock <= 0) {
        sinStock++;
      } else if (stock <= stockMin && stockMin > 0) {
        stockBajo++;
      }
      
      valorTotalInventario += (stock * precio);
    }
    
    const fechaUnMesAtras = new Date();
    fechaUnMesAtras.setMonth(fechaUnMesAtras.getMonth() - 1);
    
    let movimientosUltimoMes = 0;
    for (let i = 1; i < movimientos.length; i++) {
      if (movimientos[i][1]) {
        try {
          const fechaMov = new Date(movimientos[i][1]);
          if (fechaMov >= fechaUnMesAtras) {
            movimientosUltimoMes++;
          }
        } catch (e) {
          // Ignorar fechas inválidas
        }
      }
    }
    
    return {
      totalProductos,
      totalMovimientos,
      sinStock,
      stockBajo,
      valorTotalInventario: Math.round(valorTotalInventario * 100) / 100,
      movimientosUltimoMes
    };
  } catch (error) {
    console.error("Error en obtenerResumen:", error);
    return { totalProductos: 0, totalMovimientos: 0, sinStock: 0, stockBajo: 0, valorTotalInventario: 0 };
  }
}

function obtenerDatosAnalíticos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!hojaVentas || !hojaInventario || !hojaProductos) {
      Logger.log("⚠️ Faltan hojas requeridas para el dashboard");
      return generarDatosVacíos();
    }
    
    const datosVentas = hojaVentas.getDataRange().getValues();
    const datosInventario = hojaInventario.getDataRange().getValues();
    const datosProductos = hojaProductos.getDataRange().getValues();
    
    if (datosVentas.length <= 1) {
      Logger.log("⚠️ No hay datos de ventas");
      return generarDatosVacíos();
    }
    
    // Procesar datos
    const kpis = calcularKPIsDashboard(datosVentas, datosInventario, datosProductos);
    const ventasMensuales = calcularVentasMensuales(datosVentas, datosProductos);
    const topProductos = calcularTopProductos(datosVentas, datosProductos);
    const stockPorUbicacion = calcularStockPorUbicacion(datosInventario);
    const alertasStock = calcularAlertasStock(datosInventario, datosProductos);
    const mejoresVendedores = calcularMejoresVendedores(datosVentas);
    const topLugares = calcularTopLugares(datosVentas);
    const recomendaciones = generarRecomendaciones(kpis, alertasStock.length);
    
    return {
      kpis,
      ventasMensuales,
      topProductos,
      stockPorUbicacion,
      alertasStock,
      mejoresVendedores,
      topLugares,
      recomendaciones
    };
    
  } catch (error) {
    console.error("❌ Error en obtenerDatosAnalíticos:", error);
    Logger.log("Error stack: " + error.stack);
    return generarDatosVacíos();
  }
}

function calcularKPIsDashboard(datosVentas, datosInventario, datosProductos) {
  let ventasTotales = 0;
  let ventasMes = 0;
  let costosTotales = 0;
  let totalTransacciones = 0;
  let totalEnvios = 0;
  let totalEnviosMes = 0;
  let transaccionesConEnvio = 0; // 🆕 Contador de entregas con envío cobrado
  
  const hoy = new Date();
  const mesActual = hoy.getMonth();
  const añoActual = hoy.getFullYear();
  
  Logger.log("📅 Fecha actual: " + hoy.toLocaleDateString('es-ES'));
  Logger.log("📅 Mes actual: " + mesActual + " (0=Enero, 11=Diciembre)");
  Logger.log("📅 Año actual: " + añoActual);
  
  // Crear mapa de costos
  const costosMap = {};
  for (let i = 1; i < datosInventario.length; i++) {
    const codigo = datosInventario[i][0]?.toString().toUpperCase();
    const costo = Number(datosInventario[i][4]) || 0;
    
    if (codigo && costo > 0) {
      if (!costosMap[codigo]) {
        costosMap[codigo] = { total: 0, count: 0 };
      }
      costosMap[codigo].total += costo;
      costosMap[codigo].count++;
    }
  }
  
  const costoPromedioMap = {};
  Object.keys(costosMap).forEach(codigo => {
    costoPromedioMap[codigo] = costosMap[codigo].total / costosMap[codigo].count;
  });
  
  // ============================================
  // 🔧 SECCIÓN CORREGIDA: Procesar Ventas
  // ============================================
  Logger.log("\n📊 Procesando " + (datosVentas.length - 1) + " ventas...");
  
  for (let i = 1; i < datosVentas.length; i++) {
    try {
      const total = Number(datosVentas[i][9]) || 0;        // Columna "Total"
      const envio = Number(datosVentas[i][8]) || 0;        // Columna "Envío Cobrado"
      const fechaVenta = new Date(datosVentas[i][1]);      // Columna "Fecha"
      
      // Validar que la fecha sea válida
      if (isNaN(fechaVenta.getTime())) {
        Logger.log("⚠️ Fecha inválida en fila " + (i + 1));
        continue;
      }
      
      // Sumar totales
      ventasTotales += total;
      totalEnvios += envio;
      totalTransacciones++;
      
      if (envio > 0) {
        transaccionesConEnvio++;
      }
      
      // 🔧 CORRECCIÓN: Comparar mes y año correctamente
      const mesVenta = fechaVenta.getMonth();
      const añoVenta = fechaVenta.getFullYear();
      
      if (mesVenta === mesActual && añoVenta === añoActual) {
        ventasMes += total;
        totalEnviosMes += envio;
        
        Logger.log("✅ Venta del mes actual - Fila " + (i + 1) + ": Envío $" + envio);
      }
      
      // Calcular COGS
      const itemsTexto = datosVentas[i][6] || "";
      const items = parsearItems(itemsTexto);
      
      for (const item of items) {
        const costoUnitario = costoPromedioMap[item.codigo] || 0;
        costosTotales += (costoUnitario * item.cantidad);
      }
      
    } catch (error) {
      Logger.log("❌ Error procesando fila " + (i + 1) + ": " + error.message);
      continue;
    }
  }
  
  Logger.log("\n💰 RESUMEN DE ENVÍOS:");
  Logger.log("   Total Envíos (histórico): $" + totalEnvios.toFixed(2));
  Logger.log("   Envíos del Mes: $" + totalEnviosMes.toFixed(2));
  Logger.log("   Transacciones totales: " + totalTransacciones);
  Logger.log("   Transacciones con envío: " + transaccionesConEnvio);
  
  // Calcular Valor de Inventario
  let stockTotal = 0;
  let productosConStock = 0;
  let valorInventarioAlCosto = 0;
  
  for (let i = 1; i < datosInventario.length; i++) {
    const cantidad = Number(datosInventario[i][2]) || 0;
    const costo = Number(datosInventario[i][4]) || 0;
    
    stockTotal += cantidad;
    if (cantidad > 0) productosConStock++;
    valorInventarioAlCosto += (cantidad * costo);
  }
  
  // Calcular KPIs
  const rotacionInventario = valorInventarioAlCosto > 0 ? 
    (costosTotales / valorInventarioAlCosto) : 0;
  
  const productosUnicos = datosProductos.length - 1;
  
  const margenPromedio = ventasTotales > 0 ? 
    ((ventasTotales - costosTotales) / ventasTotales * 100) : 0;
  
  const ticketPromedio = totalTransacciones > 0 ? 
    (ventasTotales / totalTransacciones) : 0;
  
  const disponibilidad = productosUnicos > 0 ? 
    (productosConStock / productosUnicos * 100) : 0;
  
  // ============================================
  // 🔧 RETURN CORREGIDO
  // ============================================
  return {
    ventasTotales: Math.round(ventasTotales * 100) / 100,
    ventasMes: Math.round(ventasMes * 100) / 100,
    totalEnvios: Math.round(totalEnvios * 100) / 100,
    totalEnviosMes: Math.round(totalEnviosMes * 100) / 100,
    totalTransacciones: totalTransacciones, // 🆕 Agregar para cálculo de promedio
    transaccionesConEnvio: transaccionesConEnvio, // 🆕 Para estadísticas adicionales
    productosUnicos,
    stockTotal,
    margenPromedio: Math.round(margenPromedio * 10) / 10,
    rotacionInventario: Math.round(rotacionInventario * 100) / 100,
    ticketPromedio: Math.round(ticketPromedio * 100) / 100,
    disponibilidad: Math.round(disponibilidad * 10) / 10,
    _debug: {
      costosTotales: Math.round(costosTotales * 100) / 100,
      valorInventarioAlCosto: Math.round(valorInventarioAlCosto * 100) / 100,
      productosConCosto: Object.keys(costoPromedioMap).length,
      mesActual: mesActual,
      añoActual: añoActual
    }
  };
}

function calcularVentasMensuales(datosVentas, datosProductos) {
  const ventasPorMes = {};
  
  for (let i = 1; i < datosVentas.length; i++) {
    const fecha = new Date(datosVentas[i][1]);
    const mesAño = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
    const total = Number(datosVentas[i][9]) || 0;
    
    if (!ventasPorMes[mesAño]) {
      ventasPorMes[mesAño] = { ventas: 0, costos: 0 };
    }
    
    ventasPorMes[mesAño].ventas += total;
    
    // Calcular costos del mes
    const itemsTexto = datosVentas[i][6] || "";
    const items = parsearItems(itemsTexto);
    
    for (const item of items) {
      const costo = obtenerCostoProducto(item.codigo, datosProductos);
      ventasPorMes[mesAño].costos += costo * item.cantidad;
    }
  }
  
  // Convertir a array y ordenar por fecha
  const resultado = [];
  const meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  
  Object.keys(ventasPorMes).sort().slice(-6).forEach(mesAño => {
    const [año, mes] = mesAño.split('-');
    const nombreMes = meses[parseInt(mes) - 1];
    
    resultado.push({
      mes: nombreMes,
      ventas: Math.round(ventasPorMes[mesAño].ventas),
      costos: Math.round(ventasPorMes[mesAño].costos)
    });
  });
  
  return resultado;
}

function calcularTopProductos(datosVentas, datosProductos) {
  const productoStats = {};
  
  // Crear mapa de precios y nombres por código
  const productosMap = {};
  for (let i = 1; i < datosProductos.length; i++) {
    const codigo = datosProductos[i][0];
    const nombreBase = datosProductos[i][1];
    const precio = Number(datosProductos[i][5]) || 0;
    
    productosMap[codigo] = {
      nombreBase: nombreBase,
      precio: precio,
      nombreCompleto: construirNombreConVariante(codigo, nombreBase)
    };
  }
  
  // Procesar ventas
  for (let i = 1; i < datosVentas.length; i++) {
    const itemsTexto = datosVentas[i][6] || "";
    const items = parsearItems(itemsTexto);
    
    for (const item of items) {
      if (!productoStats[item.codigo]) {
        const infoProducto = productosMap[item.codigo];
        
        productoStats[item.codigo] = {
          codigo: item.codigo,
          nombreBase: infoProducto ? infoProducto.nombreBase : "Desconocido",
          nombreCompleto: infoProducto ? infoProducto.nombreCompleto : item.codigo,
          cantidad: 0,
          ingresos: 0,
          precio: infoProducto ? infoProducto.precio : 0
        };
      }
      
      productoStats[item.codigo].cantidad += item.cantidad;
      productoStats[item.codigo].ingresos += (productoStats[item.codigo].precio * item.cantidad);
    }
  }
  
  // Convertir a array y ordenar por ingresos
  const resultado = Object.values(productoStats)
    .sort((a, b) => b.ingresos - a.ingresos)
    .slice(0, 5)
    .map(p => ({
      nombre: p.nombreCompleto,        // Nombre con variante para el gráfico
      nombreCorto: acortarNombre(p.nombreCompleto, 35), // Versión corta si es muy largo
      codigo: p.codigo,                // Para referencia
      cantidad: Math.round(p.cantidad),
      ingresos: Math.round(p.ingresos)
    }));
  
  Logger.log("📊 Top 5 Productos:");
  resultado.forEach(p => {
    Logger.log(`   ${p.nombre}: ${p.cantidad} uds - $${p.ingresos}`);
  });
  
  return resultado;
}

/**
 * Construye un nombre descriptivo incluyendo la variante del código
 * Ejemplos:
 *   TSBL-XL → "Camisa de Compresión (XL)"
 *   BKBK-M → "Camisa de Bersek (M)"
 *   HUB-8 → "Hub 8 Puertos"
 */
function construirNombreConVariante(codigo, nombreBase) {
  if (!codigo || !nombreBase) {
    return codigo || nombreBase || "Desconocido";
  }
  
  // Si el código no tiene guion, devolver nombre base
  if (codigo.indexOf('-') === -1) {
    return nombreBase;
  }
  
  // Extraer la parte después del último guion
  const partes = codigo.split('-');
  const variante = partes[partes.length - 1];
  
  // Casos especiales: si la variante es solo números, podría no ser una talla
  // Ejemplo: HUB-8 → "Hub 8 Puertos" (no agregar variante)
  if (/^\d+$/.test(variante) && nombreBase.toLowerCase().includes(variante)) {
    return nombreBase;
  }
  
  // Casos de tallas comunes
  const tallasComunes = ['XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '4XL'];
  const esTalla = tallasComunes.includes(variante.toUpperCase());
  
  // Construir nombre descriptivo
  let nombreCompleto = nombreBase;
  
  // Acortar nombre base si es muy largo (más de 40 caracteres)
  if (nombreCompleto.length > 40) {
    nombreCompleto = nombreCompleto.substring(0, 40);
  }
  
  // Agregar variante entre paréntesis
  if (esTalla) {
    nombreCompleto += ` (Talla ${variante})`;
  } else if (variante.length <= 4) {
    nombreCompleto += ` (${variante})`;
  } else {
    nombreCompleto += ` - ${variante}`;
  }
  
  return nombreCompleto;
}

/**
 * Acorta un nombre si supera el límite de caracteres
 */
function acortarNombre(nombre, maxLength) {
  if (!nombre || nombre.length <= maxLength) {
    return nombre;
  }
  
  return nombre.substring(0, maxLength - 3) + '...';
}

/**
 * Parsea items vendidos (mantener función existente)
 */
function parsearItems(itemsTexto) {
  const items = [];
  
  if (!itemsTexto || itemsTexto.trim() === "") {
    return items;
  }
  
  const partes = itemsTexto.split(',');
  
  for (const parte of partes) {
    const trimmed = parte.trim();
    
    if (trimmed.includes(':')) {
      const [codigo, cantidad] = trimmed.split(':');
      
      if (codigo && cantidad) {
        items.push({
          codigo: codigo.trim().toUpperCase(),
          cantidad: parseInt(cantidad) || 0
        });
      }
    } else {
      const partesSeparadas = trimmed.split('-');
      if (partesSeparadas.length >= 3) {
        const cantidad = parseInt(partesSeparadas[partesSeparadas.length - 1]);
        const codigo = partesSeparadas.slice(0, -1).join('-');
        
        if (codigo && !isNaN(cantidad)) {
          items.push({
            codigo: codigo.trim().toUpperCase(),
            cantidad: cantidad
          });
        }
      }
    }
  }
  
  return items;
}

function calcularStockPorUbicacion(datosInventario) {
  const stockPorUbicacion = {};
  
  // Columnas: codigo, nombre, cantidad, descripcion, costo, precio, ubicacion, fecha
  for (let i = 1; i < datosInventario.length; i++) {
    const ubicacion = datosInventario[i][6] || "Sin ubicación";
    const cantidad = Number(datosInventario[i][2]) || 0;
    
    if (!stockPorUbicacion[ubicacion]) {
      stockPorUbicacion[ubicacion] = 0;
    }
    
    stockPorUbicacion[ubicacion] += cantidad;
  }
  
  return Object.keys(stockPorUbicacion).map(ub => ({
    nombre: ub,
    cantidad: Math.round(stockPorUbicacion[ub]),
    value: Math.round(stockPorUbicacion[ub])
  }));
}

function calcularAlertasStock(datosInventario, datosProductos) {
  const alertas = [];
  const stockPorCodigo = {};
  
  // Agrupar stock por código
  for (let i = 1; i < datosInventario.length; i++) {
    const codigo = datosInventario[i][0];
    const cantidad = Number(datosInventario[i][2]) || 0;
    
    if (!stockPorCodigo[codigo]) {
      stockPorCodigo[codigo] = 0;
    }
    stockPorCodigo[codigo] += cantidad;
  }
  
  // Verificar contra stock mínimo
  for (let i = 1; i < datosProductos.length; i++) {
    const codigo = datosProductos[i][0];
    const nombre = datosProductos[i][1];
    const stockMin = Number(datosProductos[i][4]) || 10;
    const stockActual = stockPorCodigo[codigo] || 0;
    
    if (stockActual <= 5 || stockActual < stockMin) {
      alertas.push({
        codigo,
        nombre,
        stock: stockActual,
        minimo: stockMin
      });
    }
  }
  
  return alertas.slice(0, 6);
}

function calcularMejoresVendedores(datosVentas) {
  const vendedorStats = {};
  
  for (let i = 1; i < datosVentas.length; i++) {
    const vendedor = datosVentas[i][4] || "Sin vendedor";
    const total = Number(datosVentas[i][9]) || 0;
    
    if (!vendedorStats[vendedor]) {
      vendedorStats[vendedor] = { ventas: 0, monto: 0 };
    }
    
    vendedorStats[vendedor].ventas++;
    vendedorStats[vendedor].monto += total;
  }
  
  return Object.keys(vendedorStats)
    .map(v => ({
      nombre: v,
      ventas: vendedorStats[v].ventas,
      monto: Math.round(vendedorStats[v].monto * 100) / 100
    }))
    .sort((a, b) => b.monto - a.monto)
    .slice(0, 3);
}

function calcularTopLugares(datosVentas) {
  const lugarStats = {};
  
  for (let i = 1; i < datosVentas.length; i++) {
    const lugar = datosVentas[i][11] || "Sin especificar";
    const total = Number(datosVentas[i][9]) || 0;
    
    if (!lugarStats[lugar]) {
      lugarStats[lugar] = { entregas: 0, monto: 0 };
    }
    
    lugarStats[lugar].entregas++;
    lugarStats[lugar].monto += total;
  }
  
  return Object.keys(lugarStats)
    .map(l => ({
      lugar: l,
      entregas: lugarStats[l].entregas,
      monto: Math.round(lugarStats[l].monto)
    }))
    .sort((a, b) => b.entregas - a.entregas)
    .slice(0, 4);
}

function generarRecomendaciones(kpis, alertasCount) {
  const recomendaciones = [];
  
  // Alertas de stock
  if (alertasCount > 0) {
    recomendaciones.push({
      tipo: 'critico',
      texto: `${alertasCount} producto(s) con stock crítico - Requieren reabastecimiento inmediato`
    });
  }
  
  // Rotación de inventario
  if (kpis.rotacionInventario > 6) {
    recomendaciones.push({
      tipo: 'advertencia',
      texto: `Rotación alta (${kpis.rotacionInventario}x) - Considera aumentar stock de productos populares`
    });
  } else if (kpis.rotacionInventario < 2) {
    recomendaciones.push({
      tipo: 'advertencia',
      texto: `Rotación baja (${kpis.rotacionInventario}x) - Evalúa productos de lenta rotación`
    });
  } else {
    recomendaciones.push({
      tipo: 'exito',
      texto: `Rotación óptima (${kpis.rotacionInventario}x) - Mantén el equilibrio actual`
    });
  }
  
  // Margen de ganancia
  if (kpis.margenPromedio >= 40) {
    recomendaciones.push({
      tipo: 'exito',
      texto: `Margen saludable (${kpis.margenPromedio}%) - Mantén la estrategia de precios`
    });
  } else if (kpis.margenPromedio < 25) {
    recomendaciones.push({
      tipo: 'critico',
      texto: `Margen bajo (${kpis.margenPromedio}%) - Revisa estructura de costos y precios`
    });
  }
  
  // Disponibilidad
  if (kpis.disponibilidad < 80) {
    recomendaciones.push({
      tipo: 'advertencia',
      texto: `Disponibilidad baja (${kpis.disponibilidad}%) - Objetivo: >95%`
    });
  } else {
    recomendaciones.push({
      tipo: 'info',
      texto: `Disponibilidad de stock al ${kpis.disponibilidad}% - Objetivo: >95%`
    });
  }
  
  return recomendaciones;
}

function exportarStockCSV() {
  try {
    const stock = obtenerStock();
    
    if (stock.length === 0) {
      return null;
    }
    
    let csv = "\uFEFF";
    csv += "Código,Nombre,Unidad,Grupo,Stock Mínimo,Stock Actual,Precio,Valor Total,Estado,Diferencia\n";
    
    stock.forEach(producto => {
      let estado = "Normal";
      let diferencia = "";
      
      if (producto.cantidad <= 0) {
        estado = "Sin Stock";
        diferencia = `-${producto.stockMin}`;
      } else if (producto.cantidad <= producto.stockMin && producto.stockMin > 0) {
        estado = "Stock Bajo";
        diferencia = `-${producto.stockMin - producto.cantidad}`;
      } else {
        diferencia = `+${producto.cantidad - producto.stockMin}`;
      }
      
      const valorTotal = producto.cantidad * producto.precio;
      
      csv += `"${producto.codigo}","${producto.nombre}","${producto.unidad}","${producto.grupo}",${producto.stockMin},${producto.cantidad},${producto.precio},${valorTotal.toFixed(2)},"${estado}","${diferencia}"\n`;
    });
    
    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const nombreArchivo = `Inventario_${fechaHora}.csv`;
    
    const blob = Utilities.newBlob(csv, 'text/csv; charset=utf-8', nombreArchivo);
    
    let carpeta;
    try {
      carpeta = DriveApp.getFoldersByName("Reportes Inventario").next();
    } catch (e) {
      carpeta = DriveApp.getRootFolder();
    }
    
    const archivo = carpeta.createFile(blob);
    
    return archivo.getUrl();
  } catch (error) {
    console.error("Error en exportarStockCSV:", error);
    return null;
  }
}

function exportarReporteConFiltros(filtros, movimientos) {
  try {
    if (!movimientos || movimientos.length === 0) {
      return null;
    }
    
    let csv = "\uFEFF"; // BOM para UTF-8
    
    // Encabezado del reporte
    csv += "=== REPORTE DE MOVIMIENTOS ===\n";
    csv += "Fecha de generación: " + new Date().toLocaleString('es-ES') + "\n";
    csv += "Período: " + filtros.fechaDesde + " al " + filtros.fechaHasta + "\n";
    
    if (filtros.tipo) csv += "Tipo de movimiento: " + getTipoMovimientoTexto(filtros.tipo) + "\n";
    if (filtros.ubicacion) csv += "Ubicación: " + filtros.ubicacion + "\n";
    if (filtros.producto) csv += "Producto: " + filtros.producto + "\n";
    if (filtros.vendedor) csv += "Vendedor: " + filtros.vendedor + "\n";
    
    csv += "Total de movimientos: " + movimientos.length + "\n\n";
    
    // Encabezados de la tabla
    csv += "Fecha,Código,Producto,Tipo,Cantidad,Ubicación,Observaciones,Usuario";
    
    // Si hay ventas en el reporte, agregar columnas adicionales
    const hayVentas = movimientos.some(m => m.tipo === 'VENTA');
    if (hayVentas) {
      csv += ",Vendedor,Entregador,Lugar Entrega,Monto Total";
    }
    csv += "\n";
    
    // Datos
    movimientos.forEach(m => {
      csv += `"${m.fecha}",`;
      csv += `"${m.codigo}",`;
      csv += `"${m.producto}",`;
      csv += `"${getTipoMovimientoTexto(m.tipo)}",`;
      csv += `"${m.cantidad}",`;
      csv += `"${m.ubicacion || '-'}",`;
      csv += `"${m.observaciones}",`;
      csv += `"${m.usuario}"`;
      
      if (hayVentas) {
        csv += `,"${m.vendedor || '-'}"`;
        csv += `,"${m.entregador || '-'}"`;
        csv += `,"${m.lugarEntrega || '-'}"`;
        csv += `,"${m.montoTotal ? '$' + m.montoTotal.toFixed(2) : '-'}"`;
      }
      
      csv += "\n";
    });
    
    // Resumen estadístico
    csv += "\n=== RESUMEN ===\n";
    
    const resumen = calcularResumenMovimientos(movimientos);
    csv += "Total de movimientos: " + resumen.total + "\n";
    
    Object.keys(resumen.porTipo).forEach(tipo => {
      csv += getTipoMovimientoTexto(tipo) + ": " + resumen.porTipo[tipo] + "\n";
    });
    
    if (resumen.totalVentas > 0) {
      csv += "\nTotal en ventas: $" + resumen.totalVentas.toFixed(2) + "\n";
    }
    
    // Crear archivo
    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const nombreArchivo = `Reporte_Movimientos_${fechaHora}.csv`;
    
    const blob = Utilities.newBlob(csv, 'text/csv; charset=utf-8', nombreArchivo);
    
    let carpeta;
    try {
      carpeta = DriveApp.getFoldersByName("Reportes Inventario").next();
    } catch (e) {
      carpeta = DriveApp.getRootFolder();
    }
    
    const archivo = carpeta.createFile(blob);
    
    return archivo.getUrl();
  } catch (error) {
    console.error("Error en exportarReporteConFiltros:", error);
    return null;
  }
}

function calcularResumenMovimientos(movimientos) {
  const resumen = {
    total: movimientos.length,
    porTipo: {},
    totalVentas: 0
  };
  
  movimientos.forEach(m => {
    // Contar por tipo
    if (!resumen.porTipo[m.tipo]) {
      resumen.porTipo[m.tipo] = 0;
    }
    resumen.porTipo[m.tipo]++;
    
    // Sumar ventas
    if (m.tipo === 'VENTA' && m.montoTotal) {
      resumen.totalVentas += m.montoTotal;
    }
  });
  
  return resumen;
}

function parsearItems(itemsTexto) {
  const items = [];
  
  if (!itemsTexto || itemsTexto.trim() === "") {
    return items;
  }
  
  // Formato esperado: "TSBL-XL:2, TSBK-M:3, HUB-8:1"
  // También soporta: "BKBK-L-2, BKBK-M-3" (separado por comas)
  const partes = itemsTexto.split(',');
  
  for (const parte of partes) {
    const trimmed = parte.trim();
    
    // Formato con dos puntos: "CODIGO:CANTIDAD"
    if (trimmed.includes(':')) {
      const [codigo, cantidad] = trimmed.split(':');
      
      if (codigo && cantidad) {
        items.push({
          codigo: codigo.trim().toUpperCase(),
          cantidad: parseInt(cantidad) || 0
        });
      }
    }
    // Formato alternativo sin dos puntos: "CODIGO-VARIANTE-CANTIDAD"
    else {
      const partesSeparadas = trimmed.split('-');
      if (partesSeparadas.length >= 3) {
        const cantidad = parseInt(partesSeparadas[partesSeparadas.length - 1]);
        const codigo = partesSeparadas.slice(0, -1).join('-');
        
        if (codigo && !isNaN(cantidad)) {
          items.push({
            codigo: codigo.trim().toUpperCase(),
            cantidad: cantidad
          });
        }
      }
    }
  }
  
  return items;
}

/**
 * Obtiene el costo promedio de un producto desde Inventario
 */
function obtenerCostoProducto(codigo, datosProductos) {
  // Esta función ahora busca en datosInventario en lugar de datosProductos
  // Se mantiene por compatibilidad pero el cálculo principal usa el mapa
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
  
  if (!hojaInventario) return 0;
  
  const datosInventario = hojaInventario.getDataRange().getValues();
  const codigoUpper = codigo.toString().toUpperCase();
  
  let totalCosto = 0;
  let count = 0;
  
  for (let i = 1; i < datosInventario.length; i++) {
    const codigoInv = datosInventario[i][0]?.toString().toUpperCase();
    
    if (codigoInv === codigoUpper) {
      const costo = Number(datosInventario[i][4]) || 0;
      if (costo > 0) {
        totalCosto += costo;
        count++;
      }
    }
  }
  
  return count > 0 ? (totalCosto / count) : 0;
}

/**
 * Obtiene el nombre de un producto
 */
function obtenerNombreProducto(codigo, datosProductos) {
  const codigoUpper = codigo.toString().toUpperCase();
  
  for (let i = 1; i < datosProductos.length; i++) {
    const codigoProd = datosProductos[i][0]?.toString().toUpperCase();
    
    if (codigoProd === codigoUpper) {
      return datosProductos[i][1] || "Desconocido";
    }
  }
  
  return "Producto no encontrado";
}

/**
 * Función de prueba para verificar el cálculo
 */
function testRotacionInventario() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaVentas = ss.getSheetByName(HOJA_VENTAS);
    const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    
    const datosVentas = hojaVentas.getDataRange().getValues();
    const datosInventario = hojaInventario.getDataRange().getValues();
    const datosProductos = hojaProductos.getDataRange().getValues();
    
    const kpis = calcularKPIsDashboard(datosVentas, datosInventario, datosProductos);
    
    Logger.log("========================================");
    Logger.log("RESULTADOS DEL TEST:");
    Logger.log("========================================");
    Logger.log("Ventas Totales: $" + kpis.ventasTotales);
    Logger.log("Rotación Inventario: " + kpis.rotacionInventario + "x");
    Logger.log("Margen Promedio: " + kpis.margenPromedio + "%");
    Logger.log("----------------------------------------");
    Logger.log("DEBUG:");
    Logger.log("COGS: $" + kpis._debug.costosTotales);
    Logger.log("Valor Inventario: $" + kpis._debug.valorInventarioAlCosto);
    Logger.log("Productos con Costo: " + kpis._debug.productosConCosto);
    Logger.log("========================================");
    
    return kpis;
  } catch (error) {
    Logger.log("❌ Error en test: " + error.message);
    Logger.log(error.stack);
  }
}