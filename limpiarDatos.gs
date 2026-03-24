/**
 * SISTEMA DE LIMPIEZA Y PROTECCIÓN DE DATOS
 * VERSIÓN SIMPLIFICADA - Sin detección de duplicados
 * 
 * La detección de duplicados se maneja con regla de validación de datos:
 * Fórmula: =CONTAR.SI($A:$A, A1)=1
 * 
 * Funciones incluidas:
 * - Limpiar datos de columnas específicas
 * - Proteger contra edición de RUTs ya ingresados
 * - Activar/desactivar protecciones manualmente
 */

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

const PROTECCION_CONFIG = {
  COLUMNA_RUT: 1,
  FILA_ENCABEZADO: 1,
  PROPERTY_KEY: "PROTECCION_RUT_ACTIVA"
};

// ============================================================================
// FUNCIONES AUXILIARES (LLAMADAS POR ORQUESTADOR)
// ============================================================================

/**
 * Verifica si la protección de RUT está activa
 * @return {boolean} - true si está activa, false si está desactivada
 */
function estaProteccionActiva() {
  const propiedades = PropertiesService.getScriptProperties();
  const estado = propiedades.getProperty(PROTECCION_CONFIG.PROPERTY_KEY);
  return estado === null || estado === "true";
}

// ============================================================================
// FUNCIONES DE ACTIVACIÓN/DESACTIVACIÓN
// ============================================================================

/**
 * DESACTIVA temporalmente la protección de edición
 * Ejecutar manualmente desde Apps Script cuando sea necesario editar RUTs
 */
function permitirEdicionRUT() {
  const propiedades = PropertiesService.getScriptProperties();
  propiedades.setProperty(PROTECCION_CONFIG.PROPERTY_KEY, "false");
  
  Logger.log("🔓 PROTECCIÓN DESACTIVADA - Ahora puedes editar RUTs libremente");
  Logger.log("⚠️ IMPORTANTE: No olvides ejecutar bloquearEdicionRUT() cuando termines");
  
  SpreadsheetApp.getUi().alert(
    "🔓 Protección Desactivada",
    "La protección de RUTs ha sido DESACTIVADA.\n\n" +
    "Ahora puedes editar RUTs existentes libremente.\n\n" +
    "⚠️ IMPORTANTE: Ejecuta 'bloquearEdicionRUT()' cuando termines de editar.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * REACTIVA la protección de edición
 * Ejecutar manualmente desde Apps Script después de permitirEdicionRUT()
 */
function bloquearEdicionRUT() {
  const propiedades = PropertiesService.getScriptProperties();
  propiedades.setProperty(PROTECCION_CONFIG.PROPERTY_KEY, "true");
  
  Logger.log("🔒 PROTECCIÓN ACTIVADA - Los RUTs están protegidos contra edición");
  
  SpreadsheetApp.getUi().alert(
    "🔒 Protección Activada",
    "La protección de RUTs ha sido ACTIVADA.\n\n" +
    "Ahora:\n" +
    "⛔ No se pueden editar RUTs existentes\n\n" +
    "Para desactivar temporalmente, ejecuta 'permitirEdicionRUT()'",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Muestra el estado actual de la protección
 * Útil para verificar si está activa o desactivada
 */
function verificarEstadoProteccion() {
  const activa = estaProteccionActiva();
  const estado = activa ? "🔒 ACTIVADA" : "🔓 DESACTIVADA";
  
  Logger.log("Estado de protección de RUT: " + estado);
  
  SpreadsheetApp.getUi().alert(
    "Estado de Protección",
    "La protección de RUTs está: " + estado + "\n\n" +
    (activa 
      ? "⛔ No se pueden editar RUTs existentes" 
      : "✓ Se pueden editar RUTs libremente"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// FUNCIONES DE LIMPIEZA
// ============================================================================

/**
 * Limpia datos de las columnas especificadas
 * VERSIÓN MEJORADA: Solicita confirmación antes de limpiar
 */
function limpiarRutCodigoYObservacion() {
  // Solicitar confirmación
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    '⚠️ Confirmar Limpieza',
    '¿Estás seguro de que deseas limpiar TODOS los datos?\n\n' +
    'Esto eliminará:\n' +
    '• RUTs (Columna A)\n' +
    '• Validaciones de RUT (Columna B)\n' +
    '• Códigos (Columna K)\n' +
    '• Validaciones de Correo (Columna H)\n' +
    '• Estados (Columna G)\n' +
    '• Observaciones (Columna M)\n' +
    '• Protecciones de celdas\n\n' +
    'Esta acción NO se puede deshacer.',
    ui.ButtonSet.YES_NO
  );
  
  // Si el usuario cancela, salir
  if (respuesta !== ui.Button.YES) {
    Logger.log("Limpieza cancelada por el usuario");
    return;
  }
  
  // Desactivar temporalmente la protección para permitir la limpieza
  const proteccionEstaba = estaProteccionActiva();
  if (proteccionEstaba) {
    PropertiesService.getScriptProperties().setProperty(PROTECCION_CONFIG.PROPERTY_KEY, "false");
    Logger.log("🔓 Protección desactivada temporalmente para limpieza");
  }
  
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ultimaFila = hoja.getLastRow();
    
    if (ultimaFila <= 1) {
      ui.alert("No hay datos para limpiar (solo encabezados)");
      return;
    }

    // Limpiar datos de la columna "RUT" (A)
    var rangoRut = hoja.getRange('A2:A' + ultimaFila);
    rangoRut.clearContent();

    // Eliminar restricciones de la columna "RUT"
    eliminarProteccionesColumnaRut(hoja);

    // Limpiar datos de la columna "VALIDACION DE RUT" (B)
    var rangoValidacion = hoja.getRange('B2:B' + ultimaFila);
    rangoValidacion.clearContent();

    // Limpiar datos de la columna "Estado" (G)
    var rangoEstado = hoja.getRange('G2:G' + ultimaFila);
    rangoEstado.clearContent();

    // Limpiar datos de la columna "VALIDACION DE CORREO" (H)
    var rangoValidacionCorreo = hoja.getRange('H2:H' + ultimaFila);
    rangoValidacionCorreo.clearContent();

    // Limpiar datos de la columna "CODIGO" (K)
    var rangoCodigo = hoja.getRange('K2:K' + ultimaFila);
    rangoCodigo.clearContent();

    // Limpiar datos de la columna "OBSERVACIONES" (M)
    var rangoObservaciones = hoja.getRange('M2:M' + ultimaFila);
    rangoObservaciones.clearContent();
    
    Logger.log("✅ Limpieza completada exitosamente");
    
    ui.alert(
      "✅ Limpieza Completada",
      "Se han limpiado exitosamente:\n" +
      "• " + (ultimaFila - 1) + " filas de datos\n" +
      "• Todas las protecciones de celdas\n\n" +
      "La hoja está lista para nuevos registros.",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("❌ Error durante la limpieza: " + error.toString());
    ui.alert("Error", "Ocurrió un error durante la limpieza: " + error.toString(), ui.ButtonSet.OK);
  } finally {
    // Reactivar la protección si estaba activa antes
    if (proteccionEstaba) {
      PropertiesService.getScriptProperties().setProperty(PROTECCION_CONFIG.PROPERTY_KEY, "true");
      Logger.log("🔒 Protección reactivada después de limpieza");
    }
  }
}

/**
 * Elimina todas las protecciones de la columna RUT
 * @param {Sheet} hoja - La hoja donde eliminar protecciones
 */
function eliminarProteccionesColumnaRut(hoja) {
  try {
    var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var eliminadas = 0;
    
    for (var i = 0; i < protecciones.length; i++) {
      var protection = protecciones[i];
      var protectedRange = protection.getRange();
      
      // Si la protección está en la columna A (RUT), eliminarla
      if (protectedRange.getColumn() === PROTECCION_CONFIG.COLUMNA_RUT) {
        protection.remove();
        eliminadas++;
      }
    }
    
    Logger.log("Protecciones eliminadas de columna RUT: " + eliminadas);
    
  } catch (error) {
    Logger.log("Error al eliminar protecciones: " + error.toString());
  }
}

// ============================================================================
// FUNCIONES DE DIAGNÓSTICO
// ============================================================================

/**
 * Cuenta cuántas celdas de RUT están protegidas
 * Útil para diagnóstico
 */
function contarProteccionesRUT() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    let contador = 0;
    
    protecciones.forEach(function(protection) {
      const rango = protection.getRange();
      if (rango.getColumn() === PROTECCION_CONFIG.COLUMNA_RUT) {
        contador++;
      }
    });
    
    Logger.log("Total de celdas RUT protegidas: " + contador);
    
    SpreadsheetApp.getUi().alert(
      "Protecciones de RUT",
      "Total de celdas protegidas en columna RUT: " + contador,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("Error en contarProteccionesRUT: " + error.toString());
  }
}
