/**
 * SISTEMA DE ASISTENCIA CON NOTIFICACIÓN HTML
 * Función que envía correo con diseño HTML estilizado cuando se edita una fila
 */
function enviarCorreoHTML() {
  // Acceder a la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Obtener los datos de la fila editada
  var rangoEditado = SpreadsheetApp.getActiveRange();
  var fila = rangoEditado.getRow();
  
  // Verificar si la fila es mayor que 1 (para evitar el encabezado)
  if (fila <= 1) {
    return; // Salir si es el encabezado
  }
  
  // Obtener datos de las columnas
  var rut = hoja.getRange(fila, 1).getValue(); // Columna A (RUT)
  var validacion = hoja.getRange(fila, 2).getValue(); // Columna B (VALIDACION DE RUT)
  var nombre = hoja.getRange(fila, 3).getValue(); // Columna C (NOMBRE)
  var correo = hoja.getRange(fila, 10).getValue(); // Columna J (CORREO)
  var asamblea = hoja.getRange(fila, 12).getValue(); // Columna L (ASAMBLEA)
  
  // VALIDACIÓN: Solo enviar si el RUT es VÁLIDO
  if (validacion !== 'VALIDO') {
    Logger.log("RUT no válido en fila " + fila + ", no se enviará correo");
    return;
  }
  
  // VALIDACIÓN: Verificar que hay correo electrónico
  if (!correo || correo.toString().trim() === '') {
    Logger.log("No hay correo en fila " + fila);
    hoja.getRange(fila, 8).setValue('ERROR: Sin correo electrónico');
    return;
  }
  
  // Verificar si ya existe un código en la columna "CODIGO"
  var codigoExistente = hoja.getRange(fila, 11).getValue(); // Columna K (CODIGO)
  var codigo = codigoExistente || generarCodigoUnico(11);
  
  // Establecer el código en la celda (no cambia si ya existe)
  hoja.getRange(fila, 11).setValue(codigo);
  
  // Proteger la celda de la columna "RUT" para evitar modificaciones
  protegerCeldaRut(hoja, fila);
  
  // Generar el HTML del correo
  var htmlBody = generarHTMLCorreo(nombre, rut, codigo, asamblea);
  
  // Configurar el asunto del correo
  var asunto = 'Confirmación de Asistencia - Asamblea Sindical ' + (asamblea || 'Sin especificar');
  
  // Versión en texto plano (fallback para clientes de correo que no soportan HTML)
  var textoPlano = 'Estimado/a ' + nombre + ',\n\n' +
                   'Se ha registrado su asistencia a la asamblea sindical del mes de ' + asamblea + '.\n\n' +
                   'DATOS DE CONFIRMACIÓN:\n' +
                   '- RUT: ' + rut + '\n' +
                   '- Código de Verificación: ' + codigo + '\n' +
                   '- Asamblea: ' + asamblea + '\n\n' +
                   'INFORMACIÓN IMPORTANTE:\n' +
                   'Si al final del mes aparece con multa en su liquidación de sueldo, puede apelar con este correo de confirmación en la página de apelación de multa del sindicato.\n\n' +
                   'Portal de Afiliados: https://www.sindicatoslim3.com/aplicaciones/app-login\n\n' +
                   'Saludos,\n' +
                   'Atte. Dpto Comunicaciones SLIM n°3';
  
  try {
    // Enviar correo con HTML
    MailApp.sendEmail({
      to: correo,
      subject: asunto,
      body: textoPlano,  // Texto plano como fallback
      htmlBody: htmlBody  // HTML como contenido principal
    });
    
    // Actualizar columna VALIDACION DE CORREO (columna H - índice 8)
    hoja.getRange(fila, 8).setValue('Asistencia enviada al correo');
    
    Logger.log("Correo HTML enviado exitosamente a: " + correo);
    
  } catch (error) {
    Logger.log("Error al enviar correo: " + error.toString());
    hoja.getRange(fila, 8).setValue('ERROR: ' + error.toString().substring(0, 50));
  }
}

/**
 * Genera el HTML estilizado para el correo de confirmación de asistencia
 * @param {string} nombre - Nombre del socio
 * @param {string} rut - RUT del socio
 * @param {string} codigo - Código único de verificación
 * @param {string} asamblea - Mes de la asamblea
 * @return {string} - HTML completo del correo
 */
function generarHTMLCorreo(nombre, rut, codigo, asamblea) {
  var html = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Confirmación de Asistencia</title>
</head>
<body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: #f4f4f4;">
    
    <!-- Contenedor principal -->
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f4f4f4;">
        <tr>
            <td style="padding: 40px 20px;">
                
                <!-- Tarjeta de contenido -->
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
                    
                    <!-- Header con gradiente -->
                    <tr>
                        <td style="background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); padding: 40px 30px; text-align: center; border-radius: 12px 12px 0 0;">
                            <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">
                                ✓ Asistencia Confirmada
                            </h1>
                            <p style="margin: 10px 0 0 0; color: #e0e7ff; font-size: 16px;">
                                Asamblea Sindical ${asamblea || 'Sin especificar'}
                            </p>
                        </td>
                    </tr>
                    
                    <!-- Contenido principal -->
                    <tr>
                        <td style="padding: 40px 30px;">
                            
                            <!-- Saludo -->
                            <p style="margin: 0 0 25px 0; color: #1f2937; font-size: 16px; line-height: 1.6;">
                                Estimado/a <strong>${nombre || 'Socio/a'}</strong>,
                            </p>
                            
                            <p style="margin: 0 0 30px 0; color: #4b5563; font-size: 15px; line-height: 1.6;">
                                Se ha registrado exitosamente su asistencia a la asamblea sindical. A continuación encontrará sus datos de confirmación:
                            </p>
                            
                            <!-- Tarjeta de datos -->
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f9fafb; border-radius: 8px; border: 1px solid #e5e7eb;">
                                <tr>
                                    <td style="padding: 25px;">
                                        
                                        <!-- RUT -->
                                        <div style="margin-bottom: 20px;">
                                            <p style="margin: 0 0 5px 0; color: #6b7280; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;">
                                                RUT
                                            </p>
                                            <p style="margin: 0; color: #111827; font-size: 18px; font-weight: 600;">
                                                ${rut || 'N/A'}
                                            </p>
                                        </div>
                                        
                                        <!-- Código de verificación -->
                                        <div style="margin-bottom: 20px;">
                                            <p style="margin: 0 0 5px 0; color: #6b7280; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;">
                                                Código de Verificación
                                            </p>
                                            <p style="margin: 0; color: #3b82f6; font-size: 24px; font-weight: 700; font-family: 'Courier New', monospace; letter-spacing: 2px;">
                                                ${codigo || 'N/A'}
                                            </p>
                                        </div>
                                        
                                        <!-- Asamblea -->
                                        <div>
                                            <p style="margin: 0 0 5px 0; color: #6b7280; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;">
                                                Asamblea
                                            </p>
                                            <p style="margin: 0; color: #111827; font-size: 18px; font-weight: 600;">
                                                ${asamblea || 'Sin especificar'}
                                            </p>
                                        </div>
                                        
                                    </td>
                                </tr>
                            </table>
                            
                            <!-- Información importante -->
                            <div style="margin-top: 30px; padding: 20px; background-color: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 6px;">
                                <p style="margin: 0 0 10px 0; color: #92400e; font-size: 14px; font-weight: 700;">
                                    ⚠️ Información Importante
                                </p>
                                <p style="margin: 0; color: #78350f; font-size: 14px; line-height: 1.5;">
                                    Si al final del mes aparece con multa en su liquidación de sueldo, puede apelar presentando este correo de confirmación en la página de apelación de multa del sindicato.
                                </p>
                            </div>
                            
                            <!-- Botón de acción -->
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-top: 35px;">
                                <tr>
                                    <td style="text-align: center;">
                                        <a href="https://www.sindicatoslim3.com/aplicaciones/app-login" style="display: inline-block; padding: 14px 32px; background-color: #3b82f6; color: #ffffff; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 15px; box-shadow: 0 2px 4px rgba(59, 130, 246, 0.3);">
                                            Ir al Portal de Afiliados
                                        </a>
                                    </td>
                                </tr>
                            </table>
                            
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="padding: 30px; background-color: #f9fafb; border-radius: 0 0 12px 12px; border-top: 1px solid #e5e7eb;">
                            <p style="margin: 0 0 15px 0; color: #6b7280; font-size: 14px; line-height: 1.6; text-align: center;">
                                Este es un correo automático, por favor no responder.
                            </p>
                            <p style="margin: 0; color: #9ca3af; font-size: 13px; text-align: center;">
                                <strong>Dpto. Comunicaciones SLIM n°3</strong><br>
                                Sindicato Libre de Trabajadores de Metro S.A.
                            </p>
                            <p style="margin: 15px 0 0 0; text-align: center;">
                                <a href="https://www.sindicatoslim3.com" style="color: #3b82f6; text-decoration: none; font-size: 13px;">
                                    www.sindicatoslim3.com
                                </a>
                            </p>
                        </td>
                    </tr>
                    
                </table>
                
            </td>
        </tr>
    </table>
    
</body>
</html>
  `;
  
  return html;
}

/**
 * Genera un código único alfanumérico
 * @param {number} length - Longitud del código a generar
 * @return {string} - Código generado
 */
function generarCodigoUnico(length) {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * Protege la celda de la columna "RUT" para evitar modificaciones
 * @param {Sheet} hoja - La hoja donde proteger
 * @param {number} fila - Número de fila a proteger
 */
function protegerCeldaRut(hoja, fila) {
  if (fila > 1) {
    var celdaRut = hoja.getRange(fila, 1); // Columna A (RUT)
    var protection = celdaRut.protect().setDescription('RUT Protegido - Fila ' + fila);
    
    // Establecer solo al usuario actual como editor (administrador)
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  }
}
