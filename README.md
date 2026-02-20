# 📦 Sistema de Control de Inventario - Comarca

Sistema web de gestión de inventario desarrollado con Google Apps Script, diseñado para controlar productos, movimientos, ventas y análisis en múltiples ubicaciones.

## 🌟 Características Principales

### 📊 Gestión de Inventario
- **Control por ubicaciones**: Manejo de stock en diferentes ubicaciones físicas
- **Registro de movimientos**: Seguimiento completo de ingresos, salidas, ventas y transferencias
- **Stock en tiempo real**: Cálculo automático de existencias por producto y ubicación
- **Transferencias entre ubicaciones**: Movimiento de productos entre diferentes almacenes
- **Importación masiva desde CSV**: Carga de inventario con distribución multi-almacén en una sola operación

### 🛍️ Gestión de Ventas
- **Registro detallado de ventas**: Información completa de vendedor, entregador, lugares y horarios
- **Canal de Venta**: Seguimiento del origen de cada venta (Facebook, WhatsApp, Instagram, TikTok, Presencial)
- **Múltiples productos por venta**: Soporte para ventas con varios ítems
- **Control de envíos**: Registro de costos de envío y lugares de entrega
- **Descuento automático de inventario**: Actualización instantánea del stock al registrar ventas

### 📈 Análisis y Reportes
- **Dashboard Analítico Completo**: Vista integral con KPIs, gráficos y métricas de rendimiento
- **KPIs de ventas**: Métricas clave (ventas totales, ticket promedio, margen de ganancia, rotación de inventario)
- **Análisis de Canales de Venta**: Seguimiento de ventas por canal (Facebook, WhatsApp, Instagram, TikTok, Presencial)
- **Reportes con filtros**: Análisis por fechas, vendedores, productos y ubicaciones
- **Visualización de datos**: Gráficos interactivos con Chart.js (líneas, barras, donut)
- **Análisis de Envíos**: Métricas detalladas de costos y lugares de entrega
- **Top Productos y Vendedores**: Rankings de mejor desempeño
- **Alertas de Stock**: Notificaciones de productos con stock crítico
- **Recomendaciones Inteligentes**: Sugerencias basadas en datos
- **Historial completo**: Trazabilidad de todos los movimientos

### 🎯 Gestión de Productos
- **Catálogo completo**: Registro de productos con código, nombre, grupo y unidad
- **Búsqueda inteligente**: Autocompletado por código o nombre
- **Organización por grupos**: Clasificación de productos por categorías
- **Unidades de medida**: Soporte para diferentes tipos de unidades

## 🏗️ Arquitectura del Sistema

### Estructura de Archivos

```
Inventory_System_appscript/
├── 📄 main.gs                          # Punto de entrada de la aplicación
├── ⚙️ config.gs                        # Configuración global y constantes
├── 🛠️ Utils.gs                         # Funciones utilitarias
│
├── 🎨 Frontend
│   ├── index.html                      # Plantilla principal HTML
│   ├── Global_CSS.html                 # Estilos globales
│   ├── Global_JS.html                  # Lógica JavaScript del cliente
│   ├── views_of_the_system.html        # Vistas del sistema
│   ├── Comp_Sidebar.html               # Componente de navegación lateral
│   ├── Comp_ModalVenta.html            # Modal para registro de ventas
│   └── Comp_ModalTransferencia.html    # Modal para transferencias
│
├── 🔧 Backend Services
│   ├── Service_Inventario.gs           # Lógica de inventario y movimientos
│   ├── Service_Productos.gs            # Gestión de productos
│   ├── Service_Ventas.gs               # Procesamiento de ventas
│   ├── Service_Analisis.gs             # Análisis y reportes
│   ├── Service_Importacion.gs          # Importación masiva desde CSV
│   └── System_Admin.gs                 # Funciones administrativas
│
└── 📋 appsscript.json                  # Configuración del proyecto Apps Script
```

### Servicios Backend

#### 📦 Service_Inventario.gs
- `insertarProductoConUbicacion()` - Registra productos con ubicación
- `registrarMovimiento()` - Registra movimientos de inventario
- `buscarEnInventarioPorUbicacion()` - Búsqueda de stock por ubicación
- `obtenerStock()` - Obtiene stock actual de todos los productos
- `calcularStock()` - Calcula stock de un producto específico
- `verificarStockEnUbicacion()` - Verifica disponibilidad en ubicación
- `descontarDeInventario()` - Descuenta stock de una ubicación
- `obtenerUbicaciones()` - Lista todas las ubicaciones disponibles
- `procesarTransferenciaEntreUbicaciones()` - Transfiere stock entre ubicaciones
- `sumarAInventario()` - Suma stock a una ubicación

#### 🏷️ Service_Productos.gs
- `registrarProducto()` - Registra nuevos productos
- `buscarProductoPorCodigo()` - Búsqueda exacta por código
- `buscarProducto()` - Búsqueda por texto
- `autocompletarProductoPorCodigo()` - Autocompletado de productos
- `obtenerProductosParaFiltro()` - Lista productos para filtros
- `obtenerListas()` - Obtiene listas de grupos y unidades

#### 💰 Service_Ventas.gs
- `registrarVentaDetallada()` - Registra ventas con detalles completos
- `obtenerReporteVentas()` - Genera reportes de ventas con filtros
- `calcularKPIsVentas()` - Calcula métricas de rendimiento
- `obtenerInfoVentaPorObservacion()` - Recupera información de ventas
- `obtenerVendedores()` - Lista vendedores únicos

#### 📊 Service_Analisis.gs
- `obtenerDatosAnalíticos()` - Genera dashboard completo con KPIs y métricas
- `calcularKPIsRendimiento()` - Calcula rotación, margen, ticket promedio, disponibilidad
- `obtenerVentasMensuales()` - Análisis de ventas vs costos por mes
- `obtenerTopProductos()` - Productos más vendidos con ingresos
- `obtenerStockPorUbicacion()` - Distribución de inventario por ubicación
- `obtenerVentasPorCanal()` - Análisis de ventas por canal de origen
- `obtenerMejoresVendedores()` - Ranking de vendedores por desempeño
- `obtenerTopLugares()` - Lugares con más entregas
- `obtenerAlertasStock()` - Productos con stock crítico
- `generarRecomendaciones()` - Sugerencias inteligentes basadas en datos

#### 📤 Service_Importacion.gs
- `importarInventarioMasivo()` - Importación masiva con sincronización completa en 3 hojas:
  - **Productos**: Crea/actualiza productos automáticamente
  - **Movimientos**: Registra ingresos por almacén
  - **Inventario**: Actualiza stock actual (suma si existe, crea si no)
- Optimizado con caché en memoria y batch operations para máximo rendimiento

## 📋 Hojas de Google Sheets

El sistema utiliza las siguientes hojas en Google Sheets:

| Hoja | Descripción |
|------|-------------|
| **Productos** | Catálogo de productos (código, nombre, grupo, unidad) |
| **Movimientos** | Registro de todos los movimientos de inventario |
| **Inventario** | Stock actual por producto y ubicación |
| **Unidades** | Tipos de unidades de medida |
| **Grupos** | Categorías de productos |
| **Entrada de Productos** | Registro de entradas al inventario |
| **Ventas** | Registro detallado de todas las ventas |

## 🚀 Instalación y Configuración

### Requisitos Previos
- Cuenta de Google
- Acceso a Google Apps Script
- Google Sheets

### Pasos de Instalación

1. **Crear una copia del Google Sheet**
   - Crea un nuevo Google Sheet o usa uno existente
   - Anota el ID del spreadsheet (se encuentra en la URL)

2. **Configurar el proyecto Apps Script**
   - Abre el editor de Apps Script desde el menú: `Extensiones > Apps Script`
   - Copia todos los archivos `.gs` y `.html` al proyecto
   - Actualiza el `SPREADSHEET_ID` en `config.gs` con tu ID de spreadsheet

3. **Configurar las hojas**
   - Crea las siguientes hojas en tu spreadsheet:
     - Productos
     - Movimientos
     - Unidades
     - Grupos
     - Inventario
     - Entrada de Productos
     - Ventas

4. **Desplegar la aplicación web**
   - En el editor de Apps Script, ve a `Implementar > Nueva implementación`
   - Selecciona tipo: `Aplicación web`
   - Configura:
     - **Ejecutar como**: Usuario que implementa
     - **Quién tiene acceso**: Según tus necesidades
   - Copia la URL de la aplicación web

5. **Configurar zona horaria**
   - Verifica que la zona horaria en `appsscript.json` sea correcta
   - Por defecto está configurada para `America/Managua`

## 🎯 Uso del Sistema

### Registro de Productos
1. Accede a la sección de "Productos"
2. Completa el formulario con código, nombre, grupo y unidad
3. Haz clic en "Registrar Producto"

### Importación Masiva desde CSV
1. Accede a la sección de "Entrada de Productos"
2. Desplázate a "Importación Masiva desde CSV"
3. Selecciona tu archivo CSV con la estructura requerida:
   - **Delimitador**: Punto y coma (;)
   - **Columna A**: Nombre del Producto
   - **Columna C**: Cantidad para Casa Dylan
   - **Columna D**: Cantidad para Casa Luden
   - **Columna E**: Cantidad para Casa Jean
   - **Columna F**: Código SKU
   - **Columna G**: Costo de Compra
   - **Columna H**: Precio de Venta
4. Haz clic en "Procesar Importación"
5. Revisa el resumen con productos creados/actualizados y distribución por almacén

**Nota**: El sistema saltará automáticamente filas de categorías (sin código SKU)

### Registro de Movimientos
1. Selecciona el tipo de movimiento (Ingreso/Salida/Transferencia)
2. Busca el producto por código o nombre
3. Especifica cantidad y ubicación
4. Agrega observaciones si es necesario
5. Confirma el registro

### Registro de Ventas
1. Abre el modal de ventas
2. Completa información del vendedor y entregador
3. **Selecciona el canal de venta** (Facebook, WhatsApp, Instagram, TikTok, Presencial)
4. Agrega productos a la venta
5. Especifica lugares de extracción y entrega
6. Registra montos y horarios
7. Confirma la venta (el stock se descuenta automáticamente)

### Transferencias entre Ubicaciones
1. Abre el modal de transferencias
2. Selecciona producto, cantidad y ubicaciones origen/destino
3. Agrega observaciones
4. Confirma la transferencia

### Consulta de Reportes y Análisis
1. Accede a la sección de "Dashboard Analítico"
2. Visualiza KPIs principales:
   - Ventas totales y del mes
   - Productos únicos y stock total
   - Rotación de inventario y margen de ganancia
   - Ticket promedio y disponibilidad
3. Revisa gráficos interactivos:
   - Ventas vs Costos (6 meses)
   - Distribución de Stock por ubicación
   - Top 5 Productos más vendidos
   - **Análisis de Canales de Venta** (nuevo)
   - Análisis de envíos y lugares
4. Consulta rankings:
   - Mejores vendedores
   - Lugares con más entregas
5. Revisa alertas de stock crítico
6. Lee recomendaciones inteligentes del sistema

## 🔧 Configuración Avanzada

### Tipos de Movimiento
```javascript
TIPOS_MOVIMIENTO = {
  INGRESO: "INGRESO",
  SALIDA: "SALIDA",
  VENTA: "VENTA",
  TRANSFERENCIA: "TRANSFERENCIA"
}
```

### Campos de Venta
```javascript
CAMPOS_VENTA = {
  VENDEDOR: "vendedor",
  ENTREGADOR: "entregador",
  CANAL: "canal", // NUEVO: Facebook, WhatsApp, Instagram, TikTok, Presencial
  ITEMS: "items",
  MONTO_COBRADO: "montoCobrado",
  LUGAR_EXTRACCION: "lugarExtraccion",
  LUGAR_ENTREGA: "lugarEntrega",
  ENVIO_COBRADO: "envioCobrado",
  HORA_SALIDA: "horaSalida",
  HORA_FINALIZACION: "horaFinalizacion"
}
```

### Canales de Venta Disponibles
```javascript
CANALES_VENTA = [
  "Facebook Marketplace",
  "WhatsApp",
  "Instagram",
  "TikTok",
  "Presencial/Local"
]
```

## 🎨 Características de la Interfaz

- **Diseño responsive**: Adaptable a dispositivos móviles y tablets
- **Layout dinámico**: Ajuste automático de altura al cambiar entre vistas
- **Navegación lateral**: Menú colapsable para fácil acceso
- **Modales interactivos**: Formularios emergentes para acciones rápidas
- **Gráficos dinámicos**: Visualización de datos con Chart.js (líneas, barras, donut)
- **Dashboard Analítico**: Vista completa con 10+ métricas y gráficos interactivos
- **Autocompletado**: Búsqueda inteligente de productos
- **Validación en tiempo real**: Verificación de stock antes de operaciones
- **Scroll automático**: Reseteo de posición al cambiar de pestaña
- **Colores por canal**: Identificación visual de canales de venta

## 🔒 Seguridad

- Ejecución como usuario que implementa
- Control de acceso configurable
- Validación de datos en backend
- Manejo de errores robusto
- Logging de operaciones críticas

## 📱 Compatibilidad

- ✅ Google Chrome (recomendado)
- ✅ Mozilla Firefox
- ✅ Safari
- ✅ Microsoft Edge
- ✅ Dispositivos móviles (iOS/Android)

## 🛠️ Tecnologías Utilizadas

- **Google Apps Script**: Backend y lógica del servidor
- **Google Sheets**: Base de datos
- **HTML5/CSS3**: Estructura y estilos
- **JavaScript**: Lógica del cliente
- **Chart.js**: Visualización de datos
- **Google Apps Script HTML Service**: Renderizado de vistas

## 📝 Notas Importantes

- El sistema utiliza `createTemplateFromFile()` para incluir componentes HTML
- Los movimientos se registran con timestamp automático
- El stock se calcula en tiempo real basado en movimientos
- Las transferencias crean dos movimientos (salida y entrada)
- Las ventas generan movimientos de tipo "VENTA" automáticamente

## 🤝 Contribuciones

Este es un proyecto interno. Para sugerencias o mejoras, contacta al administrador del sistema.

## 📄 Licencia

Uso interno - Comarca

## 👥 Soporte

Para soporte técnico o consultas, contacta al equipo de desarrollo.

---

**Versión**: 2.0  
**Última actualización**: Febrero 2026  
**Zona horaria**: America/Managua

## 📝 Changelog

### Versión 2.0 (Febrero 2026)

#### 🆕 Nuevas Funcionalidades
- **Canal de Venta**: Seguimiento del origen de cada venta (Facebook, WhatsApp, Instagram, TikTok, Presencial)
- **Dashboard Analítico Completo**: Vista integral con 10+ métricas y gráficos interactivos
- **Análisis de Canales**: Gráfico circular mostrando distribución de ventas por canal
- **KPIs de Rendimiento**: Rotación de inventario, margen de ganancia, ticket promedio, disponibilidad
- **Análisis de Envíos**: Métricas detalladas de costos y lugares de entrega
- **Rankings**: Mejores vendedores y lugares con más entregas
- **Alertas de Stock**: Notificaciones de productos con stock crítico
- **Recomendaciones Inteligentes**: Sugerencias automáticas basadas en datos

#### 🔧 Mejoras Técnicas
- **Layout Dinámico**: Ajuste automático de altura al cambiar entre vistas
- **Scroll Automático**: Reseteo de posición al cambiar de pestaña
- **Optimización de Renderizado**: Mejora en la carga de gráficos y datos
- **Colores por Canal**: Identificación visual consistente para cada canal de venta
- **Función syncBodyHeight Mejorada**: Permite que el contenedor se encoja dinámicamente

#### 🐛 Correcciones
- Solucionado: Pérdida de datos en Dashboard Analítico
- Solucionado: Fondo cortado al hacer scroll
- Solucionado: Espacio vacío persistente al cambiar de vistas largas a cortas
- Solucionado: Estructura HTML duplicada
- Mejorado: Manejo de estados de pestañas con `!important`

### Versión 1.0 (Inicial)
- Sistema base de gestión de inventario
- Registro de productos, movimientos y ventas
- Transferencias entre ubicaciones
- Reportes básicos

