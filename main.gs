function doGet() {
  
  try {
    // CAMBIO IMPORTANTE AQUÍ 👇
    // 1. Usamos createTemplateFromFile (crear plantilla)
    // 2. Agregamos .evaluate() para que ejecute los 'include'
    return HtmlService.createTemplateFromFile("index")
      .evaluate() 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}