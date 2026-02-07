function insertarProductoConUbicacion(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaEntrada = ss.getSheetByName(HOJA_ENTRADA);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!hojaEntrada || !hojaInv || !hojaProductos) {
      throw new Error("Las hojas requeridas no existen. Por favor, inicialice el sistema primero desde Configuración.");
    }

    const codigo = datos.codigo.toString().trim().toUpperCase();
    const cantidad = parseFloat(datos.cantidad);
    const costo = parseFloat(datos.costo) || 0;
    const precio = parseFloat(datos.precio);
    const almacen = datos.almacen.toString().trim();
    const nombre = datos.nombre.toString().trim();
    const unidad = datos.unidad || "Unidades";
    const grupo = datos.grupo || "General";
    const stockMin = Math.max(0, parseInt(datos.stockMin) || 0);
    const descripcion = datos.descripcion ? datos.descripcion.toString().trim() : "";
    const fechaHora = new Date();

    // Validaciones
    if (!codigo || !nombre || !almacen || cantidad <= 0 || precio <= 0) {
      return "❌ Código, nombre, cantidad, precio y almacén son obligatorios.";
    }

    Logger.log(`📦 Insertando producto: ${codigo} - ${nombre} - Cantidad: ${cantidad} - Ubicación: ${almacen}`);

    // Verificar si existe el producto en Productos
    let productoExiste = false;
    const datosProductos = hojaProductos.getDataRange().getValues();
    
    for (let i = 1; i < datosProductos.length; i++) {
      if (datosProductos[i][0] && datosProductos[i][0].toString().trim().toUpperCase() === codigo) {
        productoExiste = true;
        Logger.log(`✅ Producto ${codigo} ya existe en Productos`);
        break;
      }
    }

    // Si no existe, registrarlo automáticamente
    if (!productoExiste) {
      Logger.log(`➕ Registrando nuevo producto en Productos: ${codigo}`);
      const resultado = registrarProducto({
        codigo: codigo,
        nombre: nombre,
        unidad: unidad,
        grupo: grupo,
        stockMin: stockMin,
        precio: precio
      });
      
      if (!resultado.includes("correctamente")) {
        return resultado;
      }
    }

    // === ACTUALIZAR HOJA ENTRADA ===
    Logger.log("📝 Actualizando Hoja Entrada...");
    const headersEntrada = hojaEntrada.getRange(14, 1, 1, hojaEntrada.getLastColumn()).getValues()[0];
    const colEntrada = {};
    headersEntrada.forEach((h, i) => colEntrada[h.trim().toLowerCase()] = i + 1);

    const datosEntrada = hojaEntrada.getDataRange().getValues();
    let filaEntradaExistente = null;

    for (let i = 14; i < datosEntrada.length; i++) {
      const codEntrada = datosEntrada[i][colEntrada['codigo unico del producto'] - 1]?.toString().trim();
      if (codEntrada === codigo) {
        filaEntradaExistente = i + 1;
        break;
      }
    }

    if (filaEntradaExistente) {
      const cantidadActual = Number(hojaEntrada.getRange(filaEntradaExistente, colEntrada['cantidad de entrada del producto']).getValue()) || 0;
      hojaEntrada.getRange(filaEntradaExistente, colEntrada['cantidad de entrada del producto']).setValue(cantidadActual + cantidad);
      hojaEntrada.getRange(filaEntradaExistente, colEntrada['precio']).setValue(precio);
      hojaEntrada.getRange(filaEntradaExistente, colEntrada['costo']).setValue(costo);
      hojaEntrada.getRange(filaEntradaExistente, colEntrada['fecha y hora']).setValue(fechaHora);
      Logger.log(`✅ Actualizada entrada existente en fila ${filaEntradaExistente}`);
    } else {
      const nuevaFila = hojaEntrada.getLastRow() + 1;
      hojaEntrada.getRange(nuevaFila, colEntrada['codigo unico del producto']).setValue(codigo);
      hojaEntrada.getRange(nuevaFila, colEntrada['nombre del producto']).setValue(nombre);
      hojaEntrada.getRange(nuevaFila, colEntrada['cantidad de entrada del producto']).setValue(cantidad);
      hojaEntrada.getRange(nuevaFila, colEntrada['descripción del producto']).setValue(descripcion);
      hojaEntrada.getRange(nuevaFila, colEntrada['costo']).setValue(costo);
      hojaEntrada.getRange(nuevaFila, colEntrada['precio']).setValue(precio);
      hojaEntrada.getRange(nuevaFila, colEntrada['fecha y hora']).setValue(fechaHora);
      Logger.log(`✅ Nueva entrada creada en fila ${nuevaFila}`);
    }

    // === ACTUALIZAR INVENTARIO POR UBICACIÓN ===
    Logger.log("📍 Actualizando Inventario por ubicación...");
    const headersInv = hojaInv.getRange(1, 1, 1, hojaInv.getLastColumn()).getValues()[0];
    const colInv = {};
    headersInv.forEach((h, i) => colInv[h.trim().toLowerCase()] = i + 1);
    
    Logger.log("Columnas Inventario: " + JSON.stringify(colInv));

    // Verificar que existan las columnas necesarias
    if (!colInv['ubicacion del producto']) {
      throw new Error("❌ La hoja Inventario no tiene la columna 'ubicacion del producto'. Reinicialice el sistema.");
    }

    const datosInv = hojaInv.getDataRange().getValues();
    let filaInvExistente = null;

    for (let i = 1; i < datosInv.length; i++) {
      const codInv = datosInv[i][colInv['codigo unico del producto'] - 1]?.toString().trim();
      const ubInv = datosInv[i][colInv['ubicacion del producto'] - 1]?.toString().trim();

      if (codInv === codigo && ubInv === almacen) {
        filaInvExistente = i + 1;
        Logger.log(`✅ Encontrada ubicación existente: Fila ${filaInvExistente} - ${almacen}`);
        break;
      }
    }

    if (filaInvExistente) {
      const stockActual = Number(hojaInv.getRange(filaInvExistente, colInv['cantidad de entrada del producto']).getValue()) || 0;
      hojaInv.getRange(filaInvExistente, colInv['cantidad de entrada del producto']).setValue(stockActual + cantidad);
      hojaInv.getRange(filaInvExistente, colInv['precio']).setValue(precio);
      hojaInv.getRange(filaInvExistente, colInv['costo']).setValue(costo);
      hojaInv.getRange(filaInvExistente, colInv['fecha y hora']).setValue(fechaHora);
      Logger.log(`✅ Stock actualizado en ${almacen}: ${stockActual} + ${cantidad} = ${stockActual + cantidad}`);
    } else {
      const nuevaFilaInv = hojaInv.getLastRow() + 1;
      hojaInv.getRange(nuevaFilaInv, colInv['codigo unico del producto']).setValue(codigo);
      hojaInv.getRange(nuevaFilaInv, colInv['nombre del producto']).setValue(nombre);
      hojaInv.getRange(nuevaFilaInv, colInv['cantidad de entrada del producto']).setValue(cantidad);
      hojaInv.getRange(nuevaFilaInv, colInv['precio']).setValue(precio);
      hojaInv.getRange(nuevaFilaInv, colInv['costo']).setValue(costo);
      hojaInv.getRange(nuevaFilaInv, colInv['ubicacion del producto']).setValue(almacen);
      hojaInv.getRange(nuevaFilaInv, colInv['descripción del producto']).setValue(descripcion);
      hojaInv.getRange(nuevaFilaInv, colInv['fecha y hora']).setValue(fechaHora);
      Logger.log(`✅ Nueva ubicación creada en fila ${nuevaFilaInv}: ${almacen} con ${cantidad} unidades`);
    }

    // === REGISTRAR MOVIMIENTO ===
    Logger.log("📊 Registrando movimiento...");
    registrarMovimiento({
      codigo: codigo,
      fecha: Utilities.formatDate(fechaHora, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      tipo: TIPOS_MOVIMIENTO.INGRESO,
      cantidad: cantidad,
      ubicacion: almacen,
      observaciones: `Ingreso a ${almacen}. ${descripcion}`
    });

    Logger.log(`✅ Proceso completado exitosamente para ${codigo} en ${almacen}`);
    return `✅ Producto ${productoExiste ? 'actualizado' : 'creado y registrado'} correctamente.\n📍 Ubicación: ${almacen}\n📦 Cantidad: ${cantidad} ${unidad}`;

  } catch (error) {
    console.error("❌ Error en insertarProductoConUbicacion:", error);
    Logger.log("Error stack: " + error.stack);
    return `❌ Error: ${error.message}`;
  }
}

function registrarMovimiento(mov) {
  try {
    if (!mov || !mov.codigo || !mov.fecha || !mov.tipo || !mov.cantidad) {
      return "Datos del movimiento incompletos.";
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    if (!prodSheet || !movSheet) {
      throw new Error("Las hojas del sistema no existen. Inicialice el sistema primero.");
    }
    
    if (!movSheet.getLastRow()) {
      movSheet.getRange(1, 1, 1, 9).setValues([["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante", "Ubicación"]]);
      movSheet.getRange(1, 1, 1, 9).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    const codigoNormalizado = mov.codigo.toString().trim().toUpperCase();
    const cantidad = parseFloat(mov.cantidad);
    const tipo = mov.tipo.toString().toUpperCase();
    
    if (cantidad <= 0) {
      return "La cantidad debe ser mayor a 0.";
    }
    
    if (!Object.values(TIPOS_MOVIMIENTO).includes(tipo)) {
      return `Tipo de movimiento inválido: ${tipo}`;
    }
    
    const productos = prodSheet.getDataRange().getValues();
    let productoExiste = false;
    
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0] && productos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        productoExiste = true;
        break;
      }
    }
    
    if (!productoExiste) {
      return "El producto no existe. Regístrelo primero.";
    }
    
    const stockActual = calcularStock(codigoNormalizado);
    
    if ((tipo === TIPOS_MOVIMIENTO.SALIDA || tipo === TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO || tipo === TIPOS_MOVIMIENTO.VENTA) && stockActual < cantidad) {
      return `Stock insuficiente. Disponible: ${stockActual}, Solicitado: ${cantidad}`;
    }
    
    let stockResultante = stockActual;
    switch (tipo) {
      case TIPOS_MOVIMIENTO.INGRESO:
      case TIPOS_MOVIMIENTO.AJUSTE_POSITIVO:
        stockResultante += cantidad;
        break;
      case TIPOS_MOVIMIENTO.SALIDA:
      case TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO:
      case TIPOS_MOVIMIENTO.VENTA:
        stockResultante -= cantidad;
        break;
      case TIPOS_MOVIMIENTO.AJUSTE:
        stockResultante += cantidad;
        break;
    }
    
    stockResultante = Math.max(0, stockResultante);
    
    let fechaMovimiento;
    if (typeof mov.fecha === 'string') {
      const partesFecha = mov.fecha.split('-');
      fechaMovimiento = new Date(parseInt(partesFecha[0]), parseInt(partesFecha[1]) - 1, parseInt(partesFecha[2]), 12, 0, 0);
    } else {
      fechaMovimiento = new Date(mov.fecha);
    }
    
    movSheet.appendRow([
      codigoNormalizado,
      fechaMovimiento,
      tipo,
      cantidad,
      Session.getActiveUser().getEmail() || "Sistema",
      new Date(),
      mov.observaciones || "",
      stockResultante,
      mov.ubicacion || ""
    ]);
    
    return "Movimiento registrado correctamente.";
  } catch (error) {
    console.error("Error en registrarMovimiento:", error);
    return `Error al registrar movimiento: ${error.message}`;
  }
}

function buscarEnInventarioPorUbicacion(codigo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    
    if (!hojaInv) {
      Logger.log("❌ La hoja Inventario no existe");
      return { 
        success: false,
        error: "La hoja de inventario no existe. Inicialice el sistema primero." 
      };
    }

    const datos = hojaInv.getDataRange().getValues();
    
    if (datos.length <= 1) {
      Logger.log("⚠️ La hoja Inventario está vacía");
      return { 
        success: false,
        codigo: codigo,
        coincidencias: [], 
        totalCantidad: 0, 
        totalUbicaciones: 0,
        mensaje: "⚠️ La hoja de inventario está vacía. Primero debe registrar productos con entrada."
      };
    }

    const headers = datos[0];
    const colIndex = {};
    
    headers.forEach((h, i) => {
      if (h) {
        const headerNorm = h.toString().toLowerCase().trim();
        colIndex[headerNorm] = i;
      }
    });
    
    Logger.log("📋 Encabezados encontrados en Inventario: " + JSON.stringify(colIndex));
    
    const columnasRequeridas = ['codigo unico del producto', 'cantidad de entrada del producto', 'ubicacion del producto'];
    const columnasFaltantes = columnasRequeridas.filter(col => colIndex[col] === undefined);
    
    if (columnasFaltantes.length > 0) {
      Logger.log(`❌ Faltan columnas: ${columnasFaltantes.join(', ')}`);
      return { 
        success: false,
        error: `❌ Faltan columnas en la hoja Inventario: ${columnasFaltantes.join(', ')}. Por favor, reinicialice el sistema.` 
      };
    }
    
    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    const coincidencias = [];
    let totalCantidad = 0;

    Logger.log(`🔍 Buscando código: ${codigoNormalizado} en ${datos.length - 1} filas`);

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      
      const codigoCol = colIndex['codigo unico del producto'];
      const codigoInventario = fila[codigoCol] ? fila[codigoCol].toString().trim().toUpperCase() : '';
      
      if (codigoInventario === codigoNormalizado) {
        const cantidadCol = colIndex['cantidad de entrada del producto'];
        const nombreCol = colIndex['nombre del producto'] !== undefined ? colIndex['nombre del producto'] : 1;
        const descripcionCol = colIndex['descripción del producto'] !== undefined ? colIndex['descripción del producto'] : 3;
        const costoCol = colIndex['costo'] !== undefined ? colIndex['costo'] : 4;
        const precioCol = colIndex['precio'] !== undefined ? colIndex['precio'] : 5;
        const ubicacionCol = colIndex['ubicacion del producto'];
        const fechaCol = colIndex['fecha y hora'] !== undefined ? colIndex['fecha y hora'] : 7;
        
        const cantidad = Number(fila[cantidadCol]) || 0;
        const ubicacion = fila[ubicacionCol] ? fila[ubicacionCol].toString().trim() : 'Sin ubicación';
        
        totalCantidad += cantidad;
        
        Logger.log(`✅ Encontrado en fila ${i + 1}: ${codigoInventario} - ${ubicacion} - Cantidad: ${cantidad}`);
        
        coincidencias.push({
          fila: i + 1,
          codigo: codigoInventario,
          nombre: fila[nombreCol] ? fila[nombreCol].toString() : 'Sin nombre',
          cantidad: cantidad,
          descripcion: fila[descripcionCol] ? fila[descripcionCol].toString() : 'Sin descripción',
          costo: Number(fila[costoCol]) || 0,
          precio: Number(fila[precioCol]) || 0,
          ubicacion: ubicacion,
          fechaHora: fila[fechaCol] ? fila[fechaCol].toString() : 'Sin fecha'
        });
      }
    }

    if (coincidencias.length === 0) {
      Logger.log(`⚠️ No se encontró el código ${codigoNormalizado} en inventario`);
      return {
        success: false,
        codigo: codigo,
        coincidencias: [],
        totalCantidad: 0,
        totalUbicaciones: 0,
        mensaje: `⚠️ No se encontró el producto con código "${codigo}" en el inventario. Verifique que el producto haya sido registrado con entrada.`
      };
    }

    Logger.log(`✅ Total encontrado: ${coincidencias.length} ubicaciones, ${totalCantidad} unidades`);

    // ⚠️ AQUÍ ESTÁ LA CORRECCIÓN - AGREGAR success: true
    return {
      success: true,  // ← ESTO FALTABA
      codigo: codigo,
      coincidencias: coincidencias,
      totalCantidad: totalCantidad,
      totalUbicaciones: coincidencias.length
    };

  } catch (error) {
    console.error("❌ Error en buscarEnInventarioPorUbicacion:", error);
    Logger.log("Error stack: " + error.stack);
    return { 
      success: false,
      error: `Error al buscar: ${error.message}` 
    };
  }
}

function obtenerStock() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!prodSheet) {
      throw new Error(`La hoja '${HOJA_PRODUCTOS}' no existe.`);
    }
    
    const productos = prodSheet.getDataRange().getValues();
    
    if (productos.length <= 1) {
      return [];
    }
    
    const stock = [];
    
    for (let i = 1; i < productos.length; i++) {
      const [codigo, nombre, unidad, grupo, stockMin, precio] = productos[i];
      if (codigo && nombre) {
        const cantidad = calcularStock(codigo);
        stock.push({
          codigo: codigo.toString(),
          nombre: nombre.toString(),
          unidad: unidad || "Unidades",
          grupo: grupo || "General",
          stockMin: Math.max(0, parseInt(stockMin) || 0),
          cantidad: cantidad,
          precio: parseFloat(precio) || 0
        });
      }
    }
    
    return stock.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (error) {
    console.error("Error en obtenerStock:", error);
    return [];
  }
}

function calcularStock(codigo) {
  try {
    if (!codigo) return 0;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    if (!movSheet) {
      return 0;
    }
    
    const movimientos = movSheet.getDataRange().getValues();
    let cantidad = 0;
    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    
    for (let i = 1; i < movimientos.length; i++) {
      const [cod, fecha, tipo, cant] = movimientos[i];
      if (cod && cod.toString().trim().toUpperCase() === codigoNormalizado) {
        const valor = parseFloat(cant) || 0;
        const tipoMovimiento = tipo.toString().toUpperCase();
        
        switch (tipoMovimiento) {
          case TIPOS_MOVIMIENTO.INGRESO:
          case TIPOS_MOVIMIENTO.AJUSTE_POSITIVO:
            cantidad += valor;
            break;
          case TIPOS_MOVIMIENTO.SALIDA:
          case TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO:
          case TIPOS_MOVIMIENTO.VENTA:
            cantidad -= valor;
            break;
          case TIPOS_MOVIMIENTO.AJUSTE:
            cantidad += valor;
            break;
        }
      }
    }
    
    return Math.max(0, Math.round(cantidad * 100) / 100);
  } catch (error) {
    console.error("Error en calcularStock:", error);
    return 0;
  }
}

function verificarStockEnUbicacion(codigo, ubicacion) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    
    if (!hojaInv) return 0;
    
    const datos = hojaInv.getDataRange().getValues();
    const headers = datos[0];
    const colIndex = {};
    
    headers.forEach((h, i) => {
      colIndex[h.toString().toLowerCase().trim()] = i;
    });
    
    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    const ubicacionNormalizada = ubicacion.toString().trim();
    
    for (let i = 1; i < datos.length; i++) {
      const codInv = datos[i][colIndex['codigo unico del producto']]?.toString().trim().toUpperCase();
      const ubInv = datos[i][colIndex['ubicacion del producto']]?.toString().trim();
      
      if (codInv === codigoNormalizado && ubInv === ubicacionNormalizada) {
        return Number(datos[i][colIndex['cantidad de entrada del producto']]) || 0;
      }
    }
    
    return 0;
  } catch (error) {
    console.error("Error en verificarStockEnUbicacion:", error);
    return 0;
  }
}

function descontarDeInventario(codigo, cantidad, ubicacion) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    
    if (!hojaInv) {
      return { success: false, message: "❌ La hoja de inventario no existe." };
    }
    
    const datos = hojaInv.getDataRange().getValues();
    const headers = datos[0];
    const colIndex = {};
    
    headers.forEach((h, i) => {
      colIndex[h.toString().toLowerCase().trim()] = i + 1;
    });
    
    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    const ubicacionNormalizada = ubicacion.toString().trim();
    
    for (let i = 1; i < datos.length; i++) {
      const codInv = datos[i][colIndex['codigo unico del producto'] - 1]?.toString().trim().toUpperCase();
      const ubInv = datos[i][colIndex['ubicacion del producto'] - 1]?.toString().trim();
      
      if (codInv === codigoNormalizado && ubInv === ubicacionNormalizada) {
        const fila = i + 1;
        const stockActual = Number(hojaInv.getRange(fila, colIndex['cantidad de entrada del producto']).getValue()) || 0;
        const nuevoStock = stockActual - cantidad;
        
        if (nuevoStock < 0) {
          return { success: false, message: `❌ Stock insuficiente en ${ubicacion}` };
        }
        
        hojaInv.getRange(fila, colIndex['cantidad de entrada del producto']).setValue(nuevoStock);
        hojaInv.getRange(fila, colIndex['fecha y hora']).setValue(new Date());
        
        Logger.log(`✅ Descontado ${cantidad} de ${codigo} en ${ubicacion}. Nuevo stock: ${nuevoStock}`);
        return { success: true };
      }
    }
    
    return { success: false, message: `❌ Producto ${codigo} no encontrado en ${ubicacion}` };
  } catch (error) {
    console.error("Error en descontarDeInventario:", error);
    return { success: false, message: `❌ Error: ${error.message}` };
  }
}

function obtenerUbicaciones() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    
    // Ubicaciones predeterminadas del sistema
    const ubicacionesPredeterminadas = ["Casa Luden", "Casa Jean", "Casa Dylan"];
    
    if (!hojaInv) {
      return ubicacionesPredeterminadas;
    }
    
    const datos = hojaInv.getDataRange().getValues();
    const ubicaciones = new Set();
    
    // Agregar primero las ubicaciones predeterminadas
    ubicacionesPredeterminadas.forEach(ub => ubicaciones.add(ub));
    
    if (datos.length > 1) {
      const headers = datos[0];
      const colIndex = {};
      headers.forEach((h, i) => {
        colIndex[h.toString().toLowerCase().trim()] = i;
      });
      
      const ubicacionCol = colIndex['ubicacion del producto'];
      
      if (ubicacionCol !== undefined) {
        for (let i = 1; i < datos.length; i++) {
          if (datos[i][ubicacionCol] && datos[i][ubicacionCol].toString().trim()) {
            ubicaciones.add(datos[i][ubicacionCol].toString().trim());
          }
        }
      }
    }
    
    const ubicacionesArray = Array.from(ubicaciones).sort();
    
    Logger.log(`✅ Ubicaciones disponibles: ${ubicacionesArray.join(', ')}`);
    
    return ubicacionesArray;
  } catch (error) {
    console.error("Error en obtenerUbicaciones:", error);
    return ["Casa Luden", "Casa Jean", "Casa Dylan"];
  }
}

/**
 * Procesa una transferencia de stock entre dos ubicaciones
 * @param {Object} datos - {codigo, cantidad, ubicacionOrigen, ubicacionDestino, fecha, observaciones}
 * @returns {Object} {success: boolean, message: string}
 */
function procesarTransferenciaEntreUbicaciones(datos) {
  try {
    // Validaciones básicas
    if (!datos.codigo || !datos.cantidad || !datos.ubicacionOrigen || !datos.ubicacionDestino) {
      return {
        success: false,
        message: "❌ Todos los campos son obligatorios: código, cantidad, ubicación origen y destino."
      };
    }

    const codigo = datos.codigo.toString().trim().toUpperCase();
    const cantidad = parseFloat(datos.cantidad);
    const ubicacionOrigen = datos.ubicacionOrigen.toString().trim();
    const ubicacionDestino = datos.ubicacionDestino.toString().trim();
    const observaciones = datos.observaciones ? datos.observaciones.toString().trim() : "";
    const fecha = datos.fecha || new Date();

    // Validar que las ubicaciones sean diferentes
    if (ubicacionOrigen === ubicacionDestino) {
      return {
        success: false,
        message: "❌ La ubicación de origen y destino deben ser diferentes."
      };
    }

    // Validar cantidad positiva
    if (cantidad <= 0) {
      return {
        success: false,
        message: "❌ La cantidad debe ser mayor a cero."
      };
    }

    Logger.log(`🔄 Iniciando transferencia: ${cantidad} unidades de ${codigo} desde ${ubicacionOrigen} hacia ${ubicacionDestino}`);

    // Verificar que el producto existe
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!hojaProductos) {
      return {
        success: false,
        message: "❌ La hoja de productos no existe. Inicialice el sistema primero."
      };
    }

    const datosProductos = hojaProductos.getDataRange().getValues();
    let productoExiste = false;
    let nombreProducto = "";

    for (let i = 1; i < datosProductos.length; i++) {
      if (datosProductos[i][0] && datosProductos[i][0].toString().trim().toUpperCase() === codigo) {
        productoExiste = true;
        nombreProducto = datosProductos[i][1] || "";
        break;
      }
    }

    if (!productoExiste) {
      return {
        success: false,
        message: `❌ El producto con código "${codigo}" no existe en el catálogo.`
      };
    }

    // Verificar stock en ubicación origen
    const stockOrigen = verificarStockEnUbicacion(codigo, ubicacionOrigen);
    
    if (stockOrigen < cantidad) {
      return {
        success: false,
        message: `❌ Stock insuficiente en ${ubicacionOrigen}. Disponible: ${stockOrigen}, Solicitado: ${cantidad}`
      };
    }

    Logger.log(`✅ Stock verificado en origen: ${stockOrigen} unidades disponibles`);

    // EJECUTAR TRANSFERENCIA
    
    // 1. Descontar de ubicación origen
    const resultadoDescuento = descontarDeInventario(codigo, cantidad, ubicacionOrigen);
    
    if (!resultadoDescuento.success) {
      return {
        success: false,
        message: resultadoDescuento.message
      };
    }

    Logger.log(`✅ Descontado de ${ubicacionOrigen}`);

    // 2. Sumar a ubicación destino
    const resultadoSuma = sumarAInventario(codigo, cantidad, ubicacionDestino, nombreProducto);
    
    if (!resultadoSuma.success) {
      // Rollback: devolver el stock al origen
      Logger.log("⚠️ Error al sumar a destino. Realizando rollback...");
      sumarAInventario(codigo, cantidad, ubicacionOrigen, nombreProducto);
      
      return {
        success: false,
        message: `❌ Error al transferir a ${ubicacionDestino}. Se ha revertido la operación.`
      };
    }

    Logger.log(`✅ Sumado a ${ubicacionDestino}`);

    // 3. Registrar movimiento de transferencia
    const observacionesCompletas = `Transferencia de ${ubicacionOrigen} → ${ubicacionDestino}. ${observaciones}`;
    
    const resultadoMovimiento = registrarMovimiento({
      codigo: codigo,
      fecha: typeof fecha === 'string' ? fecha : Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      tipo: TIPOS_MOVIMIENTO.TRANSFERENCIA,
      cantidad: cantidad,
      ubicacion: `${ubicacionOrigen} → ${ubicacionDestino}`,
      observaciones: observacionesCompletas
    });

    if (!resultadoMovimiento.includes("correctamente")) {
      Logger.log("⚠️ Advertencia: El movimiento no se registró correctamente, pero la transferencia física se completó.");
    }

    Logger.log(`✅ Transferencia completada exitosamente`);

    return {
      success: true,
      message: `✅ Transferencia exitosa: ${cantidad} unidades de "${nombreProducto}" movidas desde ${ubicacionOrigen} a ${ubicacionDestino}.`
    };

  } catch (error) {
    console.error("❌ Error en procesarTransferenciaEntreUbicaciones:", error);
    Logger.log("Error stack: " + error.stack);
    return {
      success: false,
      message: `❌ Error al procesar transferencia: ${error.message}`
    };
  }
}

/**
 * Suma stock a una ubicación específica en Inventario
 * Si no existe la entrada, la crea
 * @param {string} codigo - Código del producto
 * @param {number} cantidad - Cantidad a sumar
 * @param {string} ubicacion - Ubicación destino
 * @param {string} nombreProducto - Nombre del producto (para crear entrada si no existe)
 * @returns {Object} {success: boolean, message: string}
 */
function sumarAInventario(codigo, cantidad, ubicacion, nombreProducto) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    const hojaProductos = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!hojaInv) {
      return { success: false, message: "❌ La hoja de inventario no existe." };
    }

    const datos = hojaInv.getDataRange().getValues();
    const headers = datos[0];
    const colIndex = {};
    
    headers.forEach((h, i) => {
      colIndex[h.toString().toLowerCase().trim()] = i + 1;
    });

    const codigoNormalizado = codigo.toString().trim().toUpperCase();
    const ubicacionNormalizada = ubicacion.toString().trim();
    let filaEncontrada = null;

    // Buscar si ya existe la entrada para este producto en esta ubicación
    for (let i = 1; i < datos.length; i++) {
      const codInv = datos[i][colIndex['codigo unico del producto'] - 1]?.toString().trim().toUpperCase();
      const ubInv = datos[i][colIndex['ubicacion del producto'] - 1]?.toString().trim();
      
      if (codInv === codigoNormalizado && ubInv === ubicacionNormalizada) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada) {
      // Ya existe: sumar a la cantidad actual
      const stockActual = Number(hojaInv.getRange(filaEncontrada, colIndex['cantidad de entrada del producto']).getValue()) || 0;
      const nuevoStock = stockActual + cantidad;
      
      hojaInv.getRange(filaEncontrada, colIndex['cantidad de entrada del producto']).setValue(nuevoStock);
      hojaInv.getRange(filaEncontrada, colIndex['fecha y hora']).setValue(new Date());
      
      Logger.log(`✅ Sumado ${cantidad} a ${ubicacion}. Stock anterior: ${stockActual}, nuevo: ${nuevoStock}`);
      
    } else {
      // No existe: crear nueva entrada
      const nuevaFila = hojaInv.getLastRow() + 1;
      
      // Obtener datos adicionales del producto (precio, costo, etc.)
      let precio = 0, costo = 0, descripcion = "";
      const datosProductos = hojaProductos.getDataRange().getValues();
      
      for (let i = 1; i < datosProductos.length; i++) {
        if (datosProductos[i][0] && datosProductos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
          precio = Number(datosProductos[i][5]) || 0;
          break;
        }
      }
      
      hojaInv.getRange(nuevaFila, colIndex['codigo unico del producto']).setValue(codigoNormalizado);
      hojaInv.getRange(nuevaFila, colIndex['nombre del producto']).setValue(nombreProducto);
      hojaInv.getRange(nuevaFila, colIndex['cantidad de entrada del producto']).setValue(cantidad);
      hojaInv.getRange(nuevaFila, colIndex['precio']).setValue(precio);
      hojaInv.getRange(nuevaFila, colIndex['costo']).setValue(costo);
      hojaInv.getRange(nuevaFila, colIndex['ubicacion del producto']).setValue(ubicacionNormalizada);
      hojaInv.getRange(nuevaFila, colIndex['descripción del producto']).setValue(descripcion);
      hojaInv.getRange(nuevaFila, colIndex['fecha y hora']).setValue(new Date());
      
      Logger.log(`✅ Nueva entrada creada en ${ubicacion} con ${cantidad} unidades`);
    }
    
    return { success: true };
    
  } catch (error) {
    console.error("Error en sumarAInventario:", error);
    return { success: false, message: `❌ Error: ${error.message}` };
  }
}