/**
 * Servicio de Importación Masiva de Inventario
 * Realiza sincronización completa: Productos + Movimientos + Inventario
 */

/**
 * Importa inventario masivamente desde datos CSV con sincronización completa
 * @param {Array<Array<string>>} csvData - Array bidimensional con datos del CSV
 * @returns {Object} Resumen de la importación con estadísticas
 */
function importarInventarioMasivo(csvData) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    Logger.log("🚀 Iniciando importación masiva con sincronización completa...");
    
    if (!csvData || !Array.isArray(csvData) || csvData.length === 0) {
      return { success: false, error: "❌ No se recibieron datos válidos." };
    }
    
    // Mapeo de columnas (delimitador: punto y coma)
    const COL = {
      NOMBRE: 0, CANTIDAD_DYLAN: 2, CANTIDAD_LUDEN: 3, CANTIDAD_JEAN: 4,
      CODIGO: 5, COSTO: 6, PRECIO: 7
    };
    
    const ALMACENES = [
      { nombre: "Casa Dylan", col: COL.CANTIDAD_DYLAN, key: 'DYLAN' },
      { nombre: "Casa Luden", col: COL.CANTIDAD_LUDEN, key: 'LUDEN' },
      { nombre: "Casa Jean", col: COL.CANTIDAD_JEAN, key: 'JEAN' }
    ];
    
    // Estadísticas
    const stats = {
      productosCreados: 0, productosActualizados: 0,
      unidadesDylan: 0, unidadesLuden: 0, unidadesJean: 0,
      errores: [], filasProcesadas: 0
    };
    
    // === OPTIMIZACIÓN: Cargar datos en memoria ===
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const hojaMovimientos = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    if (!hojaProductos || !hojaInventario || !hojaMovimientos) {
      return { success: false, error: "❌ Faltan hojas del sistema. Inicialice primero." };
    }
    
    // Cargar productos existentes en un mapa { "SKU": fila }
    const datosProductos = hojaProductos.getDataRange().getValues();
    const mapaProductos = {};
    for (let i = 1; i < datosProductos.length; i++) {
      const sku = datosProductos[i][0] ? datosProductos[i][0].toString().trim().toUpperCase() : '';
      if (sku) mapaProductos[sku] = i + 1; // Guardar número de fila
    }
    
    // Cargar inventario existente en un mapa { "SKU|ALMACEN": fila }
    const datosInventario = hojaInventario.getDataRange().getValues();
    const headersInv = datosInventario[0];
    const colInvCodigo = headersInv.findIndex(h => h.toString().toLowerCase().includes('codigo'));
    const colInvUbicacion = headersInv.findIndex(h => h.toString().toLowerCase().includes('ubicacion'));
    const colInvCantidad = headersInv.findIndex(h => h.toString().toLowerCase().includes('cantidad'));
    const colInvNombre = headersInv.findIndex(h => h.toString().toLowerCase().includes('nombre'));
    const colInvPrecio = headersInv.findIndex(h => h.toString().toLowerCase().includes('precio'));
    const colInvCosto = headersInv.findIndex(h => h.toString().toLowerCase().includes('costo'));
    const colInvFecha = headersInv.findIndex(h => h.toString().toLowerCase().includes('fecha'));
    
    const mapaInventario = {};
    for (let i = 1; i < datosInventario.length; i++) {
      const sku = datosInventario[i][colInvCodigo] ? datosInventario[i][colInvCodigo].toString().trim().toUpperCase() : '';
      const ubicacion = datosInventario[i][colInvUbicacion] ? datosInventario[i][colInvUbicacion].toString().trim() : '';
      if (sku && ubicacion) {
        const clave = `${sku}|${ubicacion}`;
        mapaInventario[clave] = {
          fila: i + 1,
          cantidadActual: Number(datosInventario[i][colInvCantidad]) || 0
        };
      }
    }
    
    // Arrays para batch updates
    const productosNuevos = [];
    const movimientosNuevos = [];
    const inventarioNuevo = [];
    const inventarioActualizar = [];
    
    // Buscar primera fila con datos válidos
    let primeraFilaDatos = -1;
    for (let i = 0; i < csvData.length; i++) {
      if (csvData[i][COL.CODIGO] && csvData[i][COL.CODIGO].toString().trim() !== '') {
        primeraFilaDatos = i;
        Logger.log(`✅ Primera fila de datos en índice ${i}`);
        break;
      }
    }
    
    if (primeraFilaDatos === -1) {
      return { success: false, error: "❌ No se encontraron datos válidos en el CSV." };
    }
    
    // === PROCESAR CADA FILA ===
    for (let i = primeraFilaDatos; i < csvData.length; i++) {
      const fila = csvData[i];
      
      try {
        const codigo = fila[COL.CODIGO] ? fila[COL.CODIGO].toString().trim().toUpperCase() : '';
        const nombre = fila[COL.NOMBRE] ? fila[COL.NOMBRE].toString().trim() : '';
        const costo = parseFloat(fila[COL.COSTO]) || 0;
        const precio = parseFloat(fila[COL.PRECIO]) || 0;
        
        // Saltar filas de categorías (sin código SKU)
        if (!codigo || codigo === '') {
          Logger.log(`⚠️ Fila ${i + 1}: Sin código SKU. Saltando...`);
          continue;
        }
        
        if (!nombre || precio <= 0) {
          Logger.log(`⚠️ Fila ${i + 1}: Datos incompletos. Saltando...`);
          continue;
        }
        
        Logger.log(`📦 Procesando: ${codigo} - ${nombre}`);
        
        // === 1. GESTIÓN DE PRODUCTO ===
        if (!mapaProductos[codigo]) {
          // Producto nuevo: agregar a batch
          productosNuevos.push([codigo, nombre, "Unidad", "General", 5, precio, new Date()]);
          mapaProductos[codigo] = true; // Marcar como existente en memoria
          stats.productosCreados++;
          Logger.log(`  ➕ Producto nuevo: ${codigo}`);
        } else {
          // Producto existente: actualizar precio si cambió
          const filaProducto = mapaProductos[codigo];
          if (typeof filaProducto === 'number') {
            hojaProductos.getRange(filaProducto, 6).setValue(precio);
          }
          stats.productosActualizados++;
          Logger.log(`  ✏️ Producto existente: ${codigo}`);
        }
        
        // === 2. DISTRIBUCIÓN MULTI-ALMACÉN ===
        const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        const timestamp = new Date();
        
        for (const almacen of ALMACENES) {
          const cantidad = parseFloat(fila[almacen.col]) || 0;
          
          if (cantidad > 0) {
            Logger.log(`  📍 ${almacen.nombre}: ${cantidad} unidades`);
            
            // === A. REGISTRAR MOVIMIENTO ===
            const stockActual = calcularStock(codigo); // Calcular stock antes del movimiento
            movimientosNuevos.push([
              codigo,
              timestamp,
              TIPOS_MOVIMIENTO.INGRESO,
              cantidad,
              Session.getActiveUser().getEmail() || "Sistema",
              timestamp,
              `Importación masiva CSV - ${nombre}`,
              stockActual + cantidad,
              almacen.nombre
            ]);
            
            // === B. ACTUALIZAR INVENTARIO ===
            const claveInv = `${codigo}|${almacen.nombre}`;
            
            if (mapaInventario[claveInv]) {
              // Ya existe: sumar cantidad
              const info = mapaInventario[claveInv];
              const nuevaCantidad = info.cantidadActual + cantidad;
              inventarioActualizar.push({
                fila: info.fila,
                cantidad: nuevaCantidad
              });
              mapaInventario[claveInv].cantidadActual = nuevaCantidad; // Actualizar en memoria
              Logger.log(`    ✅ Inventario actualizado: ${info.cantidadActual} + ${cantidad} = ${nuevaCantidad}`);
            } else {
              // No existe: crear nueva entrada
              inventarioNuevo.push([
                codigo,
                nombre,
                cantidad,
                "",  // descripción
                costo,
                precio,
                almacen.nombre,
                timestamp
              ]);
              mapaInventario[claveInv] = { cantidadActual: cantidad }; // Agregar a memoria
              Logger.log(`    ➕ Nueva entrada en inventario: ${cantidad} unidades`);
            }
            
            // Sumar a estadísticas
            if (almacen.key === 'DYLAN') stats.unidadesDylan += cantidad;
            if (almacen.key === 'LUDEN') stats.unidadesLuden += cantidad;
            if (almacen.key === 'JEAN') stats.unidadesJean += cantidad;
          }
        }
        
        stats.filasProcesadas++;
        
      } catch (errorFila) {
        Logger.log(`❌ Error en fila ${i + 1}: ${errorFila.message}`);
        stats.errores.push(`Fila ${i + 1}: ${errorFila.message}`);
      }
    }
    
    // === ESCRIBIR TODOS LOS CAMBIOS EN BATCH ===
    Logger.log("💾 Escribiendo cambios en batch...");
    
    // Productos nuevos
    if (productosNuevos.length > 0) {
      const ultimaFilaProductos = hojaProductos.getLastRow() + 1;
      hojaProductos.getRange(ultimaFilaProductos, 1, productosNuevos.length, 7)
        .setValues(productosNuevos);
      Logger.log(`✅ ${productosNuevos.length} productos nuevos escritos`);
    }
    
    // Movimientos nuevos
    if (movimientosNuevos.length > 0) {
      const ultimaFilaMov = hojaMovimientos.getLastRow() + 1;
      hojaMovimientos.getRange(ultimaFilaMov, 1, movimientosNuevos.length, 9)
        .setValues(movimientosNuevos);
      Logger.log(`✅ ${movimientosNuevos.length} movimientos registrados`);
    }
    
    // Inventario nuevo
    if (inventarioNuevo.length > 0) {
      const ultimaFilaInv = hojaInventario.getLastRow() + 1;
      hojaInventario.getRange(ultimaFilaInv, 1, inventarioNuevo.length, 8)
        .setValues(inventarioNuevo);
      Logger.log(`✅ ${inventarioNuevo.length} nuevas entradas en inventario`);
    }
    
    // Inventario actualizar
    if (inventarioActualizar.length > 0) {
      for (const update of inventarioActualizar) {
        hojaInventario.getRange(update.fila, colInvCantidad + 1).setValue(update.cantidad);
        hojaInventario.getRange(update.fila, colInvFecha + 1).setValue(new Date());
      }
      Logger.log(`✅ ${inventarioActualizar.length} entradas de inventario actualizadas`);
    }
    
    Logger.log("✅ Importación completada");
    Logger.log(`📊 Estadísticas: ${JSON.stringify(stats)}`);
    
    return { success: true, stats: stats };
    
  } catch (error) {
    Logger.log(`❌ Error general: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return { success: false, error: `❌ Error al procesar: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}
