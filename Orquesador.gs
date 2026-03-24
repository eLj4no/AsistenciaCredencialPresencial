/**
 * ORQUESTADOR MAESTRO DEL SISTEMA DE ASISTENCIA
 * VERSIÓN CON CONTROL DE DUPLICADOS Y CURSOR
 * 
 * ORDEN DE EJECUCIÓN:
 * 1. Verificar protección de edición (¿se puede editar?)
 * 2. Verificar duplicados Y mantener cursor en su lugar
 * 3. Validar formato de RUT (¿es un RUT válido?)
 * 4. Enviar correo electrónico (solo si es válido)
 */

// ============================================================================
// CONFIGURACIÓN DEL ORQUESTADOR
// ============================================================================

const ORQUESTADOR_CONFIG = {
  COLUMNA_RUT: 1,              // Columna A
  COLUMNA_VALIDACION: 2,       // Columna B
  COLUMNA_NOMBRE: 3,           // Columna C
  COLUMNA_VALIDACION_CORREO: 8, // Columna H
  COLUMNA_CORREO: 10,          // Columna J
  COLUMNA_CODIGO: 11,          // Columna K
  COLUMNA_ASAMBLEA: 12,        // Columna L
  FILA_ENCABEZADO: 1
};

// ============================================================================
// FUNCIÓN MAESTRA - ÚNICO TRIGGER ACTIVO
// ============================================================================

/**
 * FUNCIÓN MAESTRA que orquesta todo el sistema
 * Esta es la ÚNICA función que debe tener trigger "onEdit"
 * 
 * @param {Object} e - Objeto de evento de edición
 */
function orquestadorMaestro(e) {
  try {
    // Obtener información del evento
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    // Solo procesar ediciones en la columna A (RUT) y no en el encabezado
    if (col !== ORQUESTADOR_CONFIG.COLUMNA_RUT || row <= ORQUESTADOR_CONFIG.FILA_ENCABEZADO) {
      return;
    }
    
    const nuevoValor = e.value;
    const valorAnterior = e.oldValue;
    
    Logger.log("═══════════════════════════════════════════════════════");
    Logger.log("🎯 ORQUESTADOR INICIADO - Fila: " + row);
    Logger.log("═══════════════════════════════════════════════════════");
    
    // ========================================================================
    // PASO 1: VERIFICAR PROTECCIÓN DE EDICIÓN
    // ========================================================================
    
    if (!paso1_verificarProteccionEdicion(sheet, row, nuevoValor, valorAnterior, range)) {
      Logger.log("❌ BLOQUEADO: Protección de edición activa");
      return; // DETENER TODO EL PROCESO
    }
    
    Logger.log("✅ Paso 1 completado: Sin restricciones de edición");
    
    // ========================================================================
    // PASO 2: VERIFICAR DUPLICADOS Y MANTENER CURSOR
    // ========================================================================
    
    if (!paso2_verificarDuplicadosYMantenerCursor(sheet, row, nuevoValor, range)) {
      Logger.log("❌ BLOQUEADO: RUT duplicado detectado - cursor mantenido");
      return; // DETENER TODO EL PROCESO
    }
    
    Logger.log("✅ Paso 2 completado: RUT único");
    
    // ========================================================================
    // PASO 3: VALIDAR FORMATO DE RUT
    // ========================================================================
    
    const rutValido = paso3_validarFormatoRUT(sheet, row, nuevoValor);
    
    if (!rutValido) {
      Logger.log("⚠️ Paso 3: RUT NO VÁLIDO - No se enviará correo");
      return; // DETENER - No continuar si el RUT no es válido
    }
    
    Logger.log("✅ Paso 3 completado: RUT VÁLIDO");
    
    // ========================================================================
    // PASO 4: ENVIAR CORREO ELECTRÓNICO
    // ========================================================================
    
    // Pequeña pausa para asegurar que la validación se escribió
    Utilities.sleep(300);
    
    const correoEnviado = paso4_enviarCorreo(sheet, row);
    
    if (correoEnviado) {
      Logger.log("✅ Paso 4 completado: Correo enviado exitosamente");
    } else {
      Logger.log("⚠️ Paso 4: No se pudo enviar correo (verificar datos)");
    }
    
    Logger.log("═══════════════════════════════════════════════════════");
    Logger.log("🎉 PROCESO COMPLETADO EXITOSAMENTE");
    Logger.log("═══════════════════════════════════════════════════════");
    
  } catch (error) {
    Logger.log("═══════════════════════════════════════════════════════");
    Logger.log("💥 ERROR EN ORQUESTADOR: " + error.toString());
    Logger.log("═══════════════════════════════════════════════════════");
    
    // Registrar el error en la columna de validación de correo
    try {
      e.source.getActiveSheet()
        .getRange(e.range.getRow(), ORQUESTADOR_CONFIG.COLUMNA_VALIDACION_CORREO)
        .setValue("ERROR: " + error.toString().substring(0, 50));
    } catch (e2) {
      Logger.log("No se pudo registrar el error en la hoja");
    }
  }
}

// ============================================================================
// PASO 1: VERIFICAR PROTECCIÓN DE EDICIÓN
// ============================================================================

/**
 * Verifica si se puede editar el RUT (protección contra edición de RUTs existentes)
 * 
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {number} row - Número de fila
 * @param {string} nuevoValor - Nuevo valor ingresado
 * @param {string} valorAnterior - Valor que había antes
 * @param {Range} range - Rango editado
 * @return {boolean} - true si se puede continuar, false si está bloqueado
 */
function paso1_verificarProteccionEdicion(sheet, row, nuevoValor, valorAnterior, range) {
  // Verificar si la protección está activa
  if (!estaProteccionActiva()) {
    Logger.log("  🔓 Protección desactivada - permitiendo edición");
    return true;
  }
  
  // Si hay un valor anterior, significa que se está intentando EDITAR
  if (valorAnterior && valorAnterior !== "") {
    Logger.log("  🚫 Intento de editar RUT existente: '" + valorAnterior + "'");
    
    // Restaurar el valor anterior
    range.setValue(valorAnterior);
    
    // Mostrar alerta
    SpreadsheetApp.getUi().alert(
      "⛔ Edición Bloqueada",
      "NO SE PUEDE EDITAR: El RUT ya está registrado y protegido.\n\n" +
      "Para modificarlo, ejecuta la función 'permitirEdicionRUT()' desde Apps Script.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false; // BLOQUEAR
  }
  
  return true; // PERMITIR
}

// ============================================================================
// PASO 2: VERIFICAR DUPLICADOS Y MANTENER CURSOR
// ============================================================================

/**
 * Verifica si el RUT ya existe en otra fila
 * CRÍTICO: Limpia la celda Y mantiene el cursor en la misma posición
 * 
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {number} row - Número de fila actual
 * @param {string} nuevoValor - Nuevo RUT ingresado
 * @param {Range} range - Rango editado
 * @return {boolean} - true si NO hay duplicado, false si SÍ hay duplicado
 */
function paso2_verificarDuplicadosYMantenerCursor(sheet, row, nuevoValor, range) {
  // Si la celda está vacía, permitir
  if (!nuevoValor || nuevoValor === "") {
    Logger.log("  ℹ️ Celda vacía - no verificar duplicados");
    return true;
  }
  
  // Buscar si existe el RUT en otra fila
  const rutDuplicado = verificarRUTDuplicado(sheet, nuevoValor, row);
  
  if (rutDuplicado) {
    Logger.log("  🚫 RUT duplicado encontrado en fila " + rutDuplicado.fila);
    
    // PASO CRÍTICO 1: Limpiar la celda INMEDIATAMENTE
    range.clearContent();
    
    // PASO CRÍTICO 2: Forzar flush para que se apliquen los cambios
    SpreadsheetApp.flush();
    
    // PASO CRÍTICO 3: Usar setTimeout para activar después del flush
    // Esto evita que Google Sheets mueva el cursor
    Utilities.sleep(100);
    
    // PASO CRÍTICO 4: Mantener el foco en la misma celda
    range.activate();
    
    // PASO 5: Mostrar alerta DESPUÉS de mantener el cursor
    SpreadsheetApp.getUi().alert(
      "Ocurrió un problema",
      "Este RUT ya ha sido registrado. No es necesario volver a marcarlo. Continua con el registro de asistencia.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false; // BLOQUEAR
  }
  
  Logger.log("  ✓ RUT único - no hay duplicados");
  return true; // PERMITIR
}

/**
 * Verifica si un RUT ya existe en otra fila de la hoja
 * 
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {string} rutNuevo - El RUT a verificar
 * @param {number} filaActual - La fila actual (para excluirla de la búsqueda)
 * @return {Object|null} - Objeto con {fila, rut} si encuentra duplicado, null si no
 */
function verificarRUTDuplicado(sheet, rutNuevo, filaActual) {
  try {
    const ultimaFila = sheet.getLastRow();
    
    // Si solo hay encabezado, no hay duplicados posibles
    if (ultimaFila <= ORQUESTADOR_CONFIG.FILA_ENCABEZADO) {
      return null;
    }
    
    // Obtener todos los RUTs de la columna A
    const rangoRUTs = sheet.getRange(
      ORQUESTADOR_CONFIG.FILA_ENCABEZADO + 1,
      ORQUESTADOR_CONFIG.COLUMNA_RUT,
      ultimaFila - ORQUESTADOR_CONFIG.FILA_ENCABEZADO,
      1
    );
    
    const valores = rangoRUTs.getValues();
    
    // Normalizar el RUT nuevo para comparación
    const rutNormalizado = normalizarRUT(rutNuevo);
    
    // Buscar duplicados
    for (let i = 0; i < valores.length; i++) {
      const filaComparacion = i + ORQUESTADOR_CONFIG.FILA_ENCABEZADO + 1;
      
      // Saltar la fila actual
      if (filaComparacion === filaActual) {
        continue;
      }
      
      const rutComparacion = normalizarRUT(valores[i][0]);
      
      // Si encuentra un duplicado, retornar la información
      if (rutComparacion === rutNormalizado && rutComparacion !== "") {
        return {
          fila: filaComparacion,
          rut: valores[i][0]
        };
      }
    }
    
    // No se encontró duplicado
    return null;
    
  } catch (error) {
    Logger.log("Error en verificarRUTDuplicado: " + error.toString());
    return null; // En caso de error, permitir la operación
  }
}

/**
 * Normaliza un RUT para comparación (elimina puntos, guiones, espacios y convierte a mayúsculas)
 * 
 * @param {string} rut - El RUT a normalizar
 * @return {string} - El RUT normalizado
 */
function normalizarRUT(rut) {
  if (!rut) return "";
  return rut.toString()
    .trim()
    .toUpperCase()
    .replace(/\./g, '')      // Eliminar puntos
    .replace(/-/g, '')       // Eliminar guiones
    .replace(/\s+/g, '');    // Eliminar espacios
}

// ============================================================================
// PASO 3: VALIDAR FORMATO DE RUT
// ============================================================================

/**
 * Valida el formato del RUT chileno y escribe el resultado en columna B
 * 
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {number} row - Número de fila
 * @param {string} nuevoValor - RUT a validar
 * @return {boolean} - true si el RUT es válido, false si no
 */
function paso3_validarFormatoRUT(sheet, row, nuevoValor) {
  Logger.log("  🔍 Validando formato de RUT...");
  
  let validationResult = '';
  
  // Si la celda está vacía, limpiar validación
  if (!nuevoValor || nuevoValor === "") {
    validationResult = '';
    sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION).setValue(validationResult);
    return false;
  }
  
  try {
    // Convertir a string y validar
    const rutString = String(nuevoValor).trim();
    const esValido = validateChileanRut(rutString);
    
    validationResult = esValido ? 'VALIDO' : 'NO VALIDO';
    
    Logger.log("  " + (esValido ? "✓" : "✗") + " RUT validado: " + validationResult);
    
  } catch (error) {
    validationResult = 'NO VALIDO';
    Logger.log("  ✗ Error al validar RUT: " + error.toString());
  }
  
  // Escribir resultado en columna B
  sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION).setValue(validationResult);
  
  return validationResult === 'VALIDO';
}

// ============================================================================
// PASO 4: ENVIAR CORREO ELECTRÓNICO
// ============================================================================

/**
 * Envía el correo electrónico con diseño HTML
 * Solo se ejecuta si el RUT es válido
 * 
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {number} row - Número de fila
 * @return {boolean} - true si se envió el correo, false si no
 */
function paso4_enviarCorreo(sheet, row) {
  Logger.log("  📧 Preparando envío de correo...");
  
  try {
    // Obtener datos de la fila
    const rut = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_RUT).getValue();
    const validacion = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION).getValue();
    const nombre = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_NOMBRE).getValue();
    const correo = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_CORREO).getValue();
    const asamblea = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_ASAMBLEA).getValue();
    
    // Verificar que el RUT sea válido
    if (validacion !== 'VALIDO') {
      Logger.log("  ⚠️ RUT no válido - no enviar correo");
      return false;
    }
    
    // Verificar que hay correo electrónico
    if (!correo || correo.toString().trim() === '') {
      Logger.log("  ⚠️ No hay correo electrónico en fila " + row);
      sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION_CORREO)
        .setValue('ERROR: Sin correo electrónico');
      return false;
    }
    
    // Verificar/generar código
    let codigo = sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_CODIGO).getValue();
    if (!codigo) {
      codigo = generarCodigoUnico(11);
      sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_CODIGO).setValue(codigo);
      Logger.log("  ✓ Código generado: " + codigo);
    } else {
      Logger.log("  ✓ Código existente: " + codigo);
    }
    
    // Proteger la celda de RUT
    protegerCeldaRut(sheet, row);
    
    // Generar el HTML del correo
    const htmlBody = generarHTMLCorreo(nombre, rut, codigo, asamblea);
    
    // Configurar asunto CON ÍCONO DE TICKET VERDE ✅
    const asunto = '✅ Confirmación de Asistencia - Asamblea Sindical ' + (asamblea || 'FEBRERO 2026');
    
    // Texto plano como fallback
    const textoPlano = 'Estimado/a ' + nombre + ',\n\n' +
                       'Se ha registrado su asistencia a la asamblea sindical del mes de ' + asamblea + '.\n\n' +
                       'DATOS DE CONFIRMACIÓN:\n' +
                       '- RUT: ' + rut + '\n' +
                       '- Código de Verificación: ' + codigo + '\n' +
                       '- Asamblea: ' + asamblea + '\n\n' +
                       'Saludos,\n' +
                       'Atte. Dpto Comunicaciones SLIM n°3';
    
    // Enviar correo
    MailApp.sendEmail({
      to: correo,
      subject: asunto,
      body: textoPlano,
      htmlBody: htmlBody
    });
    
    // Actualizar columna de validación de correo
    sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION_CORREO)
      .setValue('Asistencia enviada al correo');
    
    Logger.log("  ✓ Correo enviado exitosamente a: " + correo);
    
    return true;
    
  } catch (error) {
    Logger.log("  ✗ Error al enviar correo: " + error.toString());
    
    // Registrar error en la hoja
    sheet.getRange(row, ORQUESTADOR_CONFIG.COLUMNA_VALIDACION_CORREO)
      .setValue('ERROR: ' + error.toString().substring(0, 50));
    
    return false;
  }
}

// ============================================================================
// FUNCIÓN DE PRUEBA
// ============================================================================

/**
 * Función para probar el orquestador manualmente
 * Simula la edición de una celda
 */
function probarOrquestador() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fila = 4; // Cambiar al número de fila que quieres probar
  
  Logger.log("🧪 INICIANDO PRUEBA DEL ORQUESTADOR");
  Logger.log("Fila a probar: " + fila);
  
  // Simular evento de edición
  const eventoSimulado = {
    source: SpreadsheetApp.getActiveSpreadsheet(),
    range: sheet.getRange(fila, 1),
    value: sheet.getRange(fila, 1).getValue(),
    oldValue: undefined // Simular nueva entrada
  };
  
  orquestadorMaestro(eventoSimulado);
  
  Logger.log("🧪 PRUEBA COMPLETADA");
}
