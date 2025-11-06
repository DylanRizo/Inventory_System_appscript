const SPREADSHEET_ID = "1vNX5A8tjYWHfvmRKBzmp89oNWdf_gohNllPcQigVitc";
const HOJA_PRODUCTOS = "Productos";
const HOJA_MOVIMIENTOS = "Movimientos";
const HOJA_UNIDADES = "Unidades";
const HOJA_GRUPOS = "Grupos";
const HOJA_INVENTARIO = "Inventario";
const HOJA_ENTRADA = "Entrada de Productos";
const HOJA_VENTAS = "Ventas";

const TIPOS_MOVIMIENTO = {
  INGRESO: "INGRESO",
  SALIDA: "SALIDA",
  AJUSTE_POSITIVO: "AJUSTE_POSITIVO",
  AJUSTE_NEGATIVO: "AJUSTE_NEGATIVO",
  AJUSTE: "AJUSTE",
  VENTA: "VENTA"
};
const CAMPOS_VENTA = {
  VENDEDOR: "vendedor",
  ENTREGADOR: "entregador",
  ITEMS: "items",
  MONTO_COBRADO: "montoCobrado",
  LUGAR_EXTRACCION: "lugarExtraccion",
  LUGAR_ENTREGA: "lugarEntrega",
  ENVIO_COBRADO: "envioCobrado",
  HORA_SALIDA: "horaSalida",
  HORA_FINALIZACION: "horaFinalizacion"
};
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("index")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("Sistema de Control de Inventario - Comarca");
  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial; text-align: center;">
        <h2 style="color: #dc3545;">Error del Sistema</h2>
        <p>No se pudo cargar la aplicación: ${error.message}</p>
        <button onclick="window.location.reload()">Reintentar</button>
      </div>
    `);
  }
}

// ============================================
// FUNCIONES DE PRODUCTOS
// ============================================

function registrarProducto(producto) {
  try {
    if (!producto || !producto.codigo || !producto.nombre) {
      return "Datos del producto incompletos. Código y nombre son obligatorios.";
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!sheet) {
      throw new Error(`La hoja '${HOJA_PRODUCTOS}' no existe. Inicialice el sistema primero.`);
    }
    
    if (!sheet.getLastRow()) {
      sheet.getRange(1, 1, 1, 7).setValues([["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Precio", "Fecha Creación"]]);
      sheet.getRange(1, 1, 1, 7).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    const datos = sheet.getDataRange().getValues();
    const codigoNormalizado = producto.codigo.toString().trim().toUpperCase();
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        return "Ya existe un producto con este código.";
      }
    }
    
    const nombre = producto.nombre.toString().trim();
    const unidad = producto.unidad || "Unidades";
    const grupo = producto.grupo || "General";
    const stockMin = Math.max(0, parseInt(producto.stockMin) || 0);
    const precio = Math.max(0, parseFloat(producto.precio) || 0);
    
    if (nombre.length < 2) {
      return "El nombre del producto debe tener al menos 2 caracteres.";
    }
    
    sheet.appendRow([
      codigoNormalizado,
      nombre,
      unidad,
      grupo,
      stockMin,
      precio,
      new Date()
    ]);
    
    return "Producto registrado correctamente.";
  } catch (error) {
    console.error("Error en registrarProducto:", error);
    return `Error al registrar producto: ${error.message}`;
  }
}

// ============================================
// FUNCIONES DE ENTRADA DE PRODUCTOS (INTEGRADAS)
// ============================================

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

// ============================================
// FUNCIONES DE MOVIMIENTOS
// ============================================

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

// ============================================
// FUNCIONES DE BÚSQUEDA
// ============================================

function buscarProductoPorCodigo(codigo) {
  try {
    if (!codigo || codigo.trim().length < 1) {
      return [];
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!sheet) {
      return [];
    }
    
    const datos = sheet.getDataRange().getValues();
    
    if (datos.length <= 1) {
      return [];
    }
    
    const textoBusqueda = codigo.toString().toUpperCase().trim();
    const encontrados = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && fila[0].toString().toUpperCase().startsWith(textoBusqueda)) {
        encontrados.push({
          codigo: fila[0],
          nombre: fila[1],
          unidad: fila[2] || "Unidades",
          grupo: fila[3] || "General",
          precio: fila[5] || 0
        });
      }
    }
    
    return encontrados.slice(0, 10);
  } catch (error) {
    console.error("Error en buscarProductoPorCodigo:", error);
    return [];
  }
}

function buscarProducto(texto) {
  try {
    if (!texto || texto.trim().length < 1) {
      return [];
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!sheet) {
      throw new Error(`La hoja '${HOJA_PRODUCTOS}' no existe.`);
    }
    
    const datos = sheet.getDataRange().getValues();
    
    if (datos.length <= 1) {
      return [];
    }
    
    const textoBusqueda = texto.toString().toLowerCase().trim();
    const encontrados = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && (
        fila[0].toString().toLowerCase().includes(textoBusqueda) ||
        fila[1].toString().toLowerCase().includes(textoBusqueda) ||
        (fila[3] && fila[3].toString().toLowerCase().includes(textoBusqueda))
      )) {
        const stock = calcularStock(fila[0]);
        encontrados.push([
          fila[0],
          fila[1],
          fila[2],
          fila[3],
          fila[4] || 0,
          stock,
          fila[5] || 0
        ]);
      }
    }
    
    return encontrados.sort((a, b) => a[1].localeCompare(b[1]));
  } catch (error) {
    console.error("Error en buscarProducto:", error);
    return [];
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
// ============================================
// FUNCIONES DE STOCK E INVENTARIO
// ============================================

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

// ============================================
// FUNCIONES DE REPORTES Y ANÁLISIS
// ============================================

function obtenerHistorial(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!movSheet || !prodSheet) {
      throw new Error("Las hojas del sistema no existen.");
    }
    
    const movimientos = movSheet.getDataRange().getValues();
    const productos = prodSheet.getDataRange().getValues();
    
    if (movimientos.length <= 1) {
      return [];
    }
    
    const prodMap = {};
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0]) {
        prodMap[productos[i][0].toString().toUpperCase()] = productos[i][1];
      }
    }
    
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
    
    for (let i = 1; i < movimientos.length; i++) {
      const mov = movimientos[i];
      if (!mov[colMov['Código']] || !mov[colMov['Fecha']]) continue;
      
      try {
        const fechaMov = new Date(mov[colMov['Fecha']]);
        const tipoMov = mov[colMov['Tipo']] ? mov[colMov['Tipo']].toString().toUpperCase() : "";
        
        if (fechaMov >= fechaDesde && fechaMov <= fechaHasta) {
          if (!filtros.tipo || tipoMov === filtros.tipo.toUpperCase()) {
            const codigoProducto = mov[colMov['Código']].toString().toUpperCase();
            const ubicacionCol = colMov['Ubicación'] !== undefined ? colMov['Ubicación'] : 8;
            
            resultado.push({
              codigo: mov[colMov['Código']],
              fecha: formatearFecha(fechaMov),
              tipo: tipoMov,
              cantidad: parseFloat(mov[colMov['Cantidad']]) || 0,
              producto: prodMap[codigoProducto] || "Producto no encontrado",
              observaciones: mov[colMov['Observaciones']] || "",
              usuario: mov[colMov['Usuario']] || "N/A",
              ubicacion: mov[ubicacionCol] || "-"
            });
          }
        }
      } catch (dateError) {
        console.warn(`Fecha inválida en movimiento fila ${i + 1}:`, mov[colMov['Fecha']]);
        continue;
      }
    }
    
    Logger.log(`✅ Historial generado: ${resultado.length} movimientos encontrados`);
    
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

// ============================================
// FUNCIONES DE VALIDACIÓN Y CONFIGURACIÓN
// ============================================

function validarIntegridad() {
  const errores = [];
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const hojasRequeridas = [HOJA_PRODUCTOS, HOJA_MOVIMIENTOS, HOJA_UNIDADES, HOJA_GRUPOS];
    hojasRequeridas.forEach(nombreHoja => {
      if (!ss.getSheetByName(nombreHoja)) {
        errores.push(`Falta la hoja requerida: ${nombreHoja}`);
      }
    });
    
    if (errores.length > 0) {
      return { errores };
    }
    
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    
    const productos = prodSheet.getDataRange().getValues();
    const movimientos = movSheet.getDataRange().getValues();
    
    const codigosVistos = new Set();
    for (let i = 1; i < productos.length; i++) {
      if (!productos[i][0]) continue;
      
      const codigo = productos[i][0].toString().trim().toUpperCase();
      if (codigosVistos.has(codigo)) {
        errores.push(`Código de producto duplicado: ${productos[i][0]}`);
      }
      codigosVistos.add(codigo);
      
      if (!productos[i][1] || productos[i][1].toString().trim().length < 2) {
        errores.push(`Producto ${codigo} tiene nombre inválido`);
      }
      
      const stockMin = productos[i][4];
      if (stockMin && (isNaN(stockMin) || stockMin < 0)) {
        errores.push(`Producto ${codigo} tiene stock mínimo inválido: ${stockMin}`);
      }
    }
    
    const codigosProductos = new Set();
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0]) {
        codigosProductos.add(productos[i][0].toString().trim().toUpperCase());
      }
    }
    
    for (let i = 1; i < movimientos.length; i++) {
      if (!movimientos[i][0]) continue;
      
      const codigo = movimientos[i][0].toString().trim().toUpperCase();
      const tipo = movimientos[i][2] ? movimientos[i][2].toString().toUpperCase() : "";
      const cantidad = movimientos[i][3];
      
      if (!codigosProductos.has(codigo)) {
        errores.push(`Movimiento para producto inexistente: ${movimientos[i][0]} (fila ${i + 1})`);
      }
      
      if (tipo && !Object.values(TIPOS_MOVIMIENTO).includes(tipo)) {
        errores.push(`Tipo de movimiento inválido: ${tipo} (fila ${i + 1})`);
      }
      
      if (!cantidad || isNaN(cantidad) || cantidad <= 0) {
        errores.push(`Cantidad inválida en movimiento: ${cantidad} (fila ${i + 1})`);
      }
      
      if (movimientos[i][1]) {
        try {
          new Date(movimientos[i][1]);
        } catch (e) {
          errores.push(`Fecha inválida en movimiento (fila ${i + 1}): ${movimientos[i][1]}`);
        }
      }
    }
    
    for (let i = 1; i < productos.length; i++) {
      if (!productos[i][0]) continue;
      
      const codigo = productos[i][0];
      const stock = calcularStock(codigo);
      
      if (stock < 0) {
        errores.push(`Producto ${codigo} tiene stock negativo: ${stock}`);
      }
    }
    
    return { errores };
  } catch (error) {
    errores.push(`Error al validar integridad: ${error.message}`);
    return { errores };
  }
}

function obtenerListas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let unidadesSheet = ss.getSheetByName(HOJA_UNIDADES);
    let gruposSheet = ss.getSheetByName(HOJA_GRUPOS);
    
    if (!unidadesSheet) {
      unidadesSheet = ss.insertSheet(HOJA_UNIDADES);
      const unidadesPredeterminadas = [
        ["Unidad"],
        ["Unidades"],
        ["Kilogramos"],
        ["Gramos"],
        ["Toneladas"],
        ["Litros"],
        ["Mililitros"],
        ["Metros"],
        ["Centímetros"],
        ["Metros Cuadrados"],
        ["Metros Cúbicos"],
        ["Piezas"],
        ["Cajas"],
        ["Paquetes"],
        ["Docenas"]
      ];
      unidadesSheet.getRange(1, 1, unidadesPredeterminadas.length, 1).setValues(unidadesPredeterminadas);
      unidadesSheet.getRange(1, 1).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    if (!gruposSheet) {
      gruposSheet = ss.insertSheet(HOJA_GRUPOS);
      const gruposPredeterminados = [
        ["Grupo"],
        ["Materia Prima"],
        ["Producto Terminado"],
        ["Producto en Proceso"],
        ["Herramientas"],
        ["Consumibles"],
        ["Repuestos"],
        ["Equipos"],
        ["Suministros"],
        ["Empaques"],
        ["Químicos"],
        ["General"]
      ];
      gruposSheet.getRange(1, 1, gruposPredeterminados.length, 1).setValues(gruposPredeterminados);
      gruposSheet.getRange(1, 1).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    const unidadesData = unidadesSheet.getDataRange().getValues();
    const gruposData = gruposSheet.getDataRange().getValues();
    
    const unidades = unidadesData.slice(1).map(r => r[0]).filter(u => u && u.toString().trim());
    const grupos = gruposData.slice(1).map(r => r[0]).filter(g => g && g.toString().trim());
    
    return { 
      unidades: unidades.sort(), 
      grupos: grupos.sort() 
    };
  } catch (error) {
    console.error("Error en obtenerListas:", error);
    return { 
      unidades: ["Unidades", "Kilogramos", "Litros", "Piezas"], 
      grupos: ["General", "Materia Prima", "Producto Terminado"] 
    };
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

function inicializarHojas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ===== HOJA PRODUCTOS =====
    let prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!prodSheet) {
      prodSheet = ss.insertSheet(HOJA_PRODUCTOS);
    }
    
    // FORZAR reinicio de Productos
    if (prodSheet.getLastRow() === 0 || prodSheet.getRange("A1").getValue() !== "Código") {
      prodSheet.clear();
      const encabezados = [["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Precio", "Fecha Creación"]];
      prodSheet.getRange(1, 1, 1, 7).setValues(encabezados);
      const headerRange = prodSheet.getRange(1, 1, 1, 7);
      headerRange.setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
      
      prodSheet.getRange("A:A").setNumberFormat("@");
      prodSheet.getRange("E:E").setNumberFormat("0");
      prodSheet.getRange("F:F").setNumberFormat("#,##0.00");
      prodSheet.getRange("G:G").setNumberFormat("dd/mm/yyyy hh:mm");
      
      prodSheet.setFrozenRows(1);
      prodSheet.autoResizeColumns(1, 7);
    }
    
    // ===== HOJA MOVIMIENTOS =====
    let movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet) {
      movSheet = ss.insertSheet(HOJA_MOVIMIENTOS);
    }
    
    // FORZAR reinicio de Movimientos
    if (movSheet.getLastRow() === 0 || movSheet.getRange("A1").getValue() !== "Código") {
      movSheet.clear();
      const encabezados = [["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante", "Ubicación"]];
      movSheet.getRange(1, 1, 1, 9).setValues(encabezados);
      const headerRange = movSheet.getRange(1, 1, 1, 9);
      headerRange.setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
      
      movSheet.getRange("A:A").setNumberFormat("@");
      movSheet.getRange("B:B").setNumberFormat("dd/mm/yyyy");
      movSheet.getRange("D:D").setNumberFormat("0.##");
      movSheet.getRange("F:F").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      movSheet.getRange("H:H").setNumberFormat("0.##");
      
      movSheet.setFrozenRows(1);
      movSheet.autoResizeColumns(1, 9);
    }
    
    // ===== HOJA ENTRADA DE PRODUCTOS =====
    let entradaSheet = ss.getSheetByName(HOJA_ENTRADA);
    if (!entradaSheet) {
      entradaSheet = ss.insertSheet(HOJA_ENTRADA);
    }
    
    // FORZAR reinicio de Entrada
    if (entradaSheet.getLastRow() < 14 || entradaSheet.getRange("A14").getValue() !== "codigo unico del producto") {
      entradaSheet.clear();
      const espacios = [];
      for (let i = 0; i < 13; i++) {
        espacios.push([""]);
      }
      entradaSheet.getRange(1, 1, 13, 1).setValues(espacios);
      
      const encabezados = [["codigo unico del producto", "nombre del producto", "cantidad de entrada del producto", "Descripción del Producto", "costo", "precio", "fecha y hora"]];
      entradaSheet.getRange(14, 1, 1, 7).setValues(encabezados);
      const headerRange = entradaSheet.getRange(14, 1, 1, 7);
      headerRange.setBackground("#28a745").setFontColor("white").setFontWeight("bold");
      
      entradaSheet.getRange("A:A").setNumberFormat("@");
      entradaSheet.getRange("C:C").setNumberFormat("0.##");
      entradaSheet.getRange("E:E").setNumberFormat("#,##0.00");
      entradaSheet.getRange("F:F").setNumberFormat("#,##0.00");
      entradaSheet.getRange("G:G").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      entradaSheet.setFrozenRows(14);
      entradaSheet.autoResizeColumns(1, 7);
    }
    
    // ===== HOJA INVENTARIO - ESTRUCTURA CORREGIDA =====
    let invSheet = ss.getSheetByName(HOJA_INVENTARIO);
    if (!invSheet) {
      invSheet = ss.insertSheet(HOJA_INVENTARIO);
    }
    
    // FORZAR reinicio COMPLETO de Inventario
    if (invSheet.getLastRow() === 0 || invSheet.getRange("A1").getValue() !== "codigo unico del producto") {
      invSheet.clear();
      const encabezados = [["codigo unico del producto", "nombre del producto", "cantidad de entrada del producto", "Descripción del Producto", "costo", "precio", "ubicacion del producto", "fecha y hora"]];
      invSheet.getRange(1, 1, 1, 8).setValues(encabezados);
      const headerRange = invSheet.getRange(1, 1, 1, 8);
      headerRange.setBackground("#17a2b8").setFontColor("white").setFontWeight("bold");
      
      invSheet.getRange("A:A").setNumberFormat("@");
      invSheet.getRange("C:C").setNumberFormat("0.##");
      invSheet.getRange("E:E").setNumberFormat("#,##0.00");
      invSheet.getRange("F:F").setNumberFormat("#,##0.00");
      invSheet.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      invSheet.setFrozenRows(1);
      invSheet.autoResizeColumns(1, 8);
      
      Logger.log("✅ Hoja Inventario reinicializada correctamente con estructura completa");
    }
    
    // ===== HOJA UNIDADES =====
    let unidadesSheet = ss.getSheetByName(HOJA_UNIDADES);
    if (!unidadesSheet) {
      unidadesSheet = ss.insertSheet(HOJA_UNIDADES);
      const unidadesPredeterminadas = [
        ["Unidad"],
        ["Unidades"],
        ["Cajas"],
        ["Kilogramos"],
        ["Gramos"],
        ["Toneladas"],
        ["Litros"],
        ["Mililitros"],
        ["Metros"],
        ["Centímetros"],
        ["Metros Cuadrados"],
        ["Metros Cúbicos"],
        ["Piezas"],
        ["Paquetes"],
        ["Docenas"]
      ];
      unidadesSheet.getRange(1, 1, unidadesPredeterminadas.length, 1).setValues(unidadesPredeterminadas);
      unidadesSheet.getRange(1, 1).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    
    // ===== HOJA GRUPOS =====
    let gruposSheet = ss.getSheetByName(HOJA_GRUPOS);
    if (!gruposSheet) {
      gruposSheet = ss.insertSheet(HOJA_GRUPOS);
      const gruposPredeterminados = [
        ["Grupo"],
        ["Consumibles"],
        ["Materia Prima"],
        ["Producto Terminado"],
        ["Producto en Proceso"],
        ["Herramientas"],
        ["Repuestos"],
        ["Equipos"],
        ["Suministros"],
        ["Empaques"],
        ["Químicos"],
        ["General"]
      ];
      gruposSheet.getRange(1, 1, gruposPredeterminados.length, 1).setValues(gruposPredeterminados);
      gruposSheet.getRange(1, 1).setBackground("#5DADE2").setFontColor("white").setFontWeight("bold");
    }
    // ===== HOJA VENTAS =====
let ventasSheet = ss.getSheetByName(HOJA_VENTAS);
if (!ventasSheet) {
  ventasSheet = ss.insertSheet(HOJA_VENTAS);
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
  ventasSheet.getRange(1, 1, 1, 14).setValues(encabezados);
  ventasSheet.getRange(1, 1, 1, 14)
    .setBackground("#dc3545")
    .setFontColor("white")
    .setFontWeight("bold");
  ventasSheet.setFrozenRows(1);
  ventasSheet.autoResizeColumns(1, 14);
}
    // Inicializar listas de unidades y grupos
    obtenerListas();
    
    return "✅ Sistema inicializado correctamente.\n\n📋 Hojas creadas/verificadas:\n✅ Productos\n✅ Movimientos\n✅ Entrada de Productos\n✅ Inventario (con estructura completa)\n✅ Unidades\n✅ Grupos\n\n⚠️ IMPORTANTE: La hoja Inventario ha sido reinicializada con la estructura correcta.";
  } catch (error) {
    console.error("Error en inicializarHojas:", error);
    return `❌ Error al inicializar sistema: ${error.message}`;
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
// ============================================
// FUNCIÓN registrar venta detallada 
// ============================================

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
    
    // Validar datos requeridos
    if (!datosVenta.items || datosVenta.items.length === 0) {
      return { success: false, message: "❌ Debe incluir al menos un producto en la venta." };
    }
    
    if (!datosVenta.vendedor || !datosVenta.montoCobrado) {
      return { success: false, message: "❌ Vendedor y monto cobrado son obligatorios." };
    }
    
    // Generar ID único para la venta
    const timestamp = new Date();
    const idVenta = `V-${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss")}`;
    
    // Procesar items vendidos
    const itemsTexto = datosVenta.items.map(item => 
      `${item.codigo}:${item.cantidad}`
    ).join(", ");
    
    const montoCobrado = parseFloat(datosVenta.montoCobrado) || 0;
    const envioCobrado = parseFloat(datosVenta.envioCobrado) || 0;
    const total = montoCobrado + envioCobrado;
    
    // Registrar en hoja de Ventas
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
    
    // Registrar movimientos de inventario para cada item
    const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);
    const resultadosMovimientos = [];
    
    for (const item of datosVenta.items) {
      const codigo = item.codigo.toString().trim().toUpperCase();
      const cantidad = parseFloat(item.cantidad);
      
      if (cantidad <= 0) {
        resultadosMovimientos.push(`⚠️ ${codigo}: cantidad inválida`);
        continue;
      }
      
      // Verificar stock disponible en la ubicación especificada
      const stockDisponible = verificarStockEnUbicacion(codigo, datosVenta.lugarExtraccion);
      
      if (stockDisponible < cantidad) {
        return { 
          success: false, 
          message: `❌ Stock insuficiente de ${codigo} en ${datosVenta.lugarExtraccion}. Disponible: ${stockDisponible}, Solicitado: ${cantidad}` 
        };
      }
      
      // Descontar del inventario
      const resultadoDescuento = descontarDeInventario(
        codigo, 
        cantidad, 
        datosVenta.lugarExtraccion
      );
      
      if (!resultadoDescuento.success) {
        return resultadoDescuento;
      }
      
      // Registrar movimiento
      const resultadoMov = registrarMovimiento({
        codigo: codigo,
        fecha: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tipo: TIPOS_MOVIMIENTO.VENTA,
        cantidad: cantidad,
        ubicacion: datosVenta.lugarExtraccion,
        observaciones: `Venta ${idVenta} - Vendedor: ${datosVenta.vendedor} - Entrega: ${datosVenta.lugarEntrega}`
      });
      
      resultadosMovimientos.push(`✅ ${codigo}: ${cantidad} unidades`);
    }
    
    return {
      success: true,
      message: `✅ Venta registrada exitosamente.\n\n📋 ID: ${idVenta}\n👤 Vendedor: ${datosVenta.vendedor}\n💰 Total: $${total.toFixed(2)}\n\n${resultadosMovimientos.join('\n')}`,
      idVenta: idVenta
    };
    
  } catch (error) {
    console.error("❌ Error en registrarVentaDetallada:", error);
    Logger.log("Error stack: " + error.stack);
    return { success: false, message: `❌ Error: ${error.message}` };
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
// ============================================
// FUNCIÓN AUTOCOMPLETADO PARA FORMULARIOS
// ============================================

function autocompletarProductoPorCodigo(codigo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaEntrada = ss.getSheetByName(HOJA_ENTRADA);
    
    if (!hojaEntrada || !codigo || codigo.toString().trim() === '') {
      return null;
    }

    const headersHist = hojaEntrada.getRange(14, 1, 1, hojaEntrada.getLastColumn()).getValues()[0];
    const colHist = {};
    headersHist.forEach((h, i) => colHist[h.trim().toLowerCase()] = i + 1);

    const datosHist = hojaEntrada.getDataRange().getValues();
    
    for (let i = 14; i < datosHist.length; i++) {
      const codHist = datosHist[i][colHist['codigo unico del producto'] - 1]?.toString().trim();
      
      if (codHist === codigo.toString().trim().toUpperCase()) {
        return {
          nombre: datosHist[i][colHist['nombre del producto'] - 1] || '',
          costo: datosHist[i][colHist['costo'] - 1] || 0,
          precio: datosHist[i][colHist['precio'] - 1] || 0,
          descripcion: datosHist[i][colHist['descripción del producto'] - 1] || ''
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error("Error en autocompletarProductoPorCodigo:", error);
    return null;
  }
}

// ============================================
// FUNCIONES PARA EL DASHBOARD ANALÍTICO
// ============================================

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
  
  const hoy = new Date();
  const mesActual = hoy.getMonth();
  const añoActual = hoy.getFullYear();
  
  // Columnas de Ventas: ID Venta, Fecha, Hora Salida, Hora Fin, Vendedor, Entregador, Items, Monto, Envío, Total...
  for (let i = 1; i < datosVentas.length; i++) {
    const total = Number(datosVentas[i][9]) || 0; // Columna "Total"
    const fecha = new Date(datosVentas[i][1]); // Columna "Fecha"
    
    ventasTotales += total;
    totalTransacciones++;
    
    if (fecha.getMonth() === mesActual && fecha.getFullYear() === añoActual) {
      ventasMes += total;
    }
    
    // Calcular costos aproximados de los items vendidos
    const itemsTexto = datosVentas[i][6] || ""; // Columna "Items Vendidos"
    const items = parsearItems(itemsTexto);
    
    for (const item of items) {
      const costo = obtenerCostoProducto(item.codigo, datosProductos);
      costosTotales += costo * item.cantidad;
    }
  }
  
  // Calcular productos únicos
  const productosUnicos = datosProductos.length - 1;
  
  // Calcular stock total
  let stockTotal = 0;
  let productosConStock = 0;
  
  for (let i = 1; i < datosInventario.length; i++) {
    const cantidad = Number(datosInventario[i][2]) || 0; // Columna "cantidad"
    stockTotal += cantidad;
    if (cantidad > 0) productosConStock++;
  }
  
  // Calcular margen promedio ponderado
  const margenPromedio = ventasTotales > 0 ? 
    ((ventasTotales - costosTotales) / ventasTotales * 100) : 0;
  
  // Calcular rotación de inventario (aproximado)
  const valorInventario = calcularValorInventario(datosInventario);
  const rotacionInventario = valorInventario > 0 ? 
    (costosTotales / valorInventario) : 0;
  
  // Calcular ticket promedio
  const ticketPromedio = totalTransacciones > 0 ? 
    (ventasTotales / totalTransacciones) : 0;
  
  // Calcular disponibilidad
  const disponibilidad = productosUnicos > 0 ? 
    (productosConStock / productosUnicos * 100) : 0;
  
  return {
    ventasTotales: Math.round(ventasTotales * 100) / 100,
    ventasMes: Math.round(ventasMes * 100) / 100,
    productosUnicos,
    stockTotal,
    margenPromedio: Math.round(margenPromedio * 10) / 10,
    rotacionInventario: Math.round(rotacionInventario * 10) / 10,
    ticketPromedio: Math.round(ticketPromedio * 100) / 100,
    disponibilidad: Math.round(disponibilidad * 10) / 10
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
  
  // Crear mapa de precios por código
  const preciosMap = {};
  for (let i = 1; i < datosProductos.length; i++) {
    const codigo = datosProductos[i][0];
    const precio = Number(datosProductos[i][5]) || 0;
    preciosMap[codigo] = precio;
  }
  
  // Procesar ventas
  for (let i = 1; i < datosVentas.length; i++) {
    const itemsTexto = datosVentas[i][6] || "";
    const items = parsearItems(itemsTexto);
    
    for (const item of items) {
      if (!productoStats[item.codigo]) {
        productoStats[item.codigo] = {
          codigo: item.codigo,
          nombre: obtenerNombreProducto(item.codigo, datosProductos),
          cantidad: 0,
          ingresos: 0
        };
      }
      
      productoStats[item.codigo].cantidad += item.cantidad;
      productoStats[item.codigo].ingresos += (preciosMap[item.codigo] || 0) * item.cantidad;
    }
  }
  
  // Convertir a array y ordenar por ingresos
  const resultado = Object.values(productoStats)
    .sort((a, b) => b.ingresos - a.ingresos)
    .slice(0, 5)
    .map(p => ({
      nombre: p.nombre.substring(0, 20),
      cantidad: Math.round(p.cantidad),
      ingresos: Math.round(p.ingresos)
    }));
  
  return resultado;
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

// Funciones auxiliares
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
