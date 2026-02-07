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

function obtenerProductosParaFiltro() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    
    if (!prodSheet || prodSheet.getLastRow() <= 1) {
      return [];
    }
    
    const datos = prodSheet.getDataRange().getValues();
    const productos = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1]) {
        productos.push({
          codigo: datos[i][0].toString(),
          nombre: datos[i][1].toString()
        });
      }
    }
    
    return productos.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (error) {
    console.error("Error en obtenerProductosParaFiltro:", error);
    return [];
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