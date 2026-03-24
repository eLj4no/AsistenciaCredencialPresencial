/**
 * Función que se ejecuta automáticamente cada vez que se edita la hoja de cálculo.
 * Valida el RUT en la columna A y escribe el resultado en la columna B.
 *
 * @param {Object} e El objeto de evento de edición que contiene información sobre la edición.
 */
function onEdit(e) {
  // Obtiene la hoja activa y el rango editado.
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // Define las columnas relevantes.
  const RUT_COLUMN = 1; // Columna A (RUT)
  const VALIDATION_COLUMN = 2; // Columna B (RUT VALIDADO)

  // Verifica si la edición ocurrió en la columna A y no es la fila del encabezado.
  if (range.getColumn() === RUT_COLUMN && range.getRow() > 1) {
    const rutValue = range.getValue(); // Obtiene el valor de la celda editada.
    let validationResult = '';

    // Si la celda está vacía, limpia la celda de validación.
    if (!rutValue) {
      validationResult = '';
    } else {
      // Intenta validar el RUT.
      try {
        // Convierte el valor a string para asegurar el procesamiento.
        const rutString = String(rutValue).trim();
        validationResult = validateChileanRut(rutString) ? 'VALIDO' : 'NO VALIDO';
      } catch (error) {
        // Si hay un error (ej. formato inesperado), marca como no válido.
        validationResult = 'NO VALIDO';
        console.error("Error al validar RUT:", rutValue, error);
      }
    }

    // Escribe el resultado de la validación en la columna B de la misma fila.
    sheet.getRange(range.getRow(), VALIDATION_COLUMN).setValue(validationResult);
  }
}

/**
 * Valida un RUT chileno sin puntos ni guion, incluyendo el dígito verificador.
 * Ejemplo de formato de entrada: "178592130"
 *
 * @param {string} rut El RUT a validar (ej. "178592130").
 * @returns {boolean} Verdadero si el RUT es válido, falso en caso contrario.
 */
function validateChileanRut(rut) {
  // Elimina cualquier caracter que no sea número o 'k'/'K'.
  rut = rut.replace(/[^0-9kK]/g, '').toUpperCase();

  // El RUT debe tener al menos 2 caracteres (número + dígito verificador).
  if (rut.length < 2) {
    return false;
  }

  // Separa el cuerpo del RUT del dígito verificador.
  const body = rut.substring(0, rut.length - 1);
  const dv = rut.slice(-1);

  // Asegura que el cuerpo sea numérico.
  if (!/^\d+$/.test(body)) {
    return false;
  }

  let sum = 0;
  let multiplier = 2;

  // Calcula la suma ponderada del cuerpo del RUT.
  for (let i = body.length - 1; i >= 0; i--) {
    sum += parseInt(body.charAt(i), 10) * multiplier;
    multiplier++;
    if (multiplier > 7) {
      multiplier = 2; // Reinicia el multiplicador después de 7.
    }
  }

  // Calcula el dígito verificador esperado.
  const expectedDvNum = 11 - (sum % 11);
  let expectedDv;

  if (expectedDvNum === 11) {
    expectedDv = '0';
  } else if (expectedDvNum === 10) {
    expectedDv = 'K';
  } else {
    expectedDv = String(expectedDvNum);
  }

  // Compara el dígito verificador calculado con el proporcionado.
  return expectedDv === dv;
}

/**
 * Función para ejecutar manualmente que valida todos los RUTs en la Columna A
 * y escribe los resultados en la Columna B.
 */
function validateAllRutsInColumnA() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define las columnas relevantes.
  const RUT_COLUMN = 1; // Columna A (RUT)
  const VALIDATION_COLUMN = 2; // Columna B (RUT VALIDADO)

  // Obtiene la última fila con contenido en la hoja.
  const lastRow = sheet.getLastRow();

  // Si no hay datos (solo encabezado), no hace nada.
  if (lastRow <= 1) {
    console.log("No hay RUTs para validar (solo encabezado o hoja vacía).");
    return;
  }

  // Obtiene todos los valores de la Columna A (excluyendo el encabezado).
  // Rango: (fila de inicio, columna de inicio, número de filas, número de columnas)
  const rutRange = sheet.getRange(2, RUT_COLUMN, lastRow - 1, 1);
  const rutValues = rutRange.getValues(); // Obtiene un array 2D de los valores.

  const validationResults = [];

  // Itera sobre cada RUT y valida.
  for (let i = 0; i < rutValues.length; i++) {
    const rutValue = rutValues[i][0]; // Accede al valor del RUT en el array 2D.
    let result = '';

    if (!rutValue) {
      result = ''; // Si la celda está vacía, el resultado también lo estará.
    } else {
      try {
        const rutString = String(rutValue).trim();
        result = validateChileanRut(rutString) ? 'VALIDO' : 'NO VALIDO';
      } catch (error) {
        result = 'NO VALIDO';
        console.error("Error al validar RUT en fila", i + 2, ":", rutValue, error);
      }
    }
    validationResults.push([result]); // Añade el resultado como un array 2D para setValue.
  }

  // Escribe todos los resultados de validación en la Columna B en una sola operación.
  // Rango: (fila de inicio, columna de inicio, número de filas, número de columnas)
  const validationRange = sheet.getRange(2, VALIDATION_COLUMN, validationResults.length, 1);
  validationRange.setValues(validationResults);

  console.log("Validación masiva de RUTs completada.");
}
