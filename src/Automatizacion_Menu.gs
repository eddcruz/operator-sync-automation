/**
 * @function onOpen
 * Función que se ejecuta automáticamente al abrir el archivo de Google Sheets.
 * Crea un menú personalizado llamado "Centro de Procesos" en la barra de herramientas.
 * Desde este menú se pueden ejecutar acciones automatizadas relacionadas con operadores.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Obtiene la interfaz de usuario del documento

  // Crea el menú personalizado con distintas acciones disponibles
  ui.createMenu('Automatización')
    .addItem('Importar operadores (correo)', 'procesarCorreosOperadores') 
    // Extrae archivos semanales con operadores y actualiza la base de datos
    .addItem('Consultas de faltantes de operador (correo)', 'revisarFaltantesOperador') 
    // Busca y responde correos de descuentos de operadores
    .addToUi(); // Agrega el menú a la interfaz de usuario
}
