// ══════════════════════════════════════════════
//  MENÚ PERSONALIZADO
// ══════════════════════════════════════════════
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('💱 Monitor Divisas')
      .addItem('🔄 Actualizar tasas ahora', 'actualizarTasas')
      .addItem('📊 Actualizar Dashboard', 'actualizarDashboard')
      .addItem('📧 Generar y enviar reporte', 'generarYEnviarReporte')
      .addSeparator()
      .addItem('🗑️ Limpiar tasas de hoy', 'limpiarTasasHoy')
      .addToUi();
  } catch (e) {
    Logger.log('Error en onOpen: ' + e.message);
  }
}


// ══════════════════════════════════════════════
//  OBTENER TASAS DESDE LA API
// ══════════════════════════════════════════════
function obtenerTasasAPI() {
  const url = `https://v6.exchangerate-api.com/v6/${CONFIG.API_KEY}/latest/USD`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  if (data.result !== 'success') {
    throw new Error('Error en la API: ' + data['error-type']);
  }

  return data.conversion_rates;
}


// ══════════════════════════════════════════════
//  ACTUALIZAR TASAS EN EL SHEETS
// ══════════════════════════════════════════════
function actualizarTasas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHoy = ss.getSheetByName('Tasas-Hoy');
  const hojaHistorial = ss.getSheetByName('Historial');

  try {
    const tasas = obtenerTasasAPI();
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

    // ── Guardar tasas de ayer antes de actualizar
    const tasasAyer = {};
    const ultimaFilaHistorial = hojaHistorial.getLastRow();
    if (ultimaFilaHistorial >= 2) {
      const filaAyer = hojaHistorial.getRange(ultimaFilaHistorial, 1, 1, 6).getValues()[0];
      tasasAyer['COP'] = filaAyer[1];
      tasasAyer['EUR'] = filaAyer[2];
      tasasAyer['MXN'] = filaAyer[3];
      tasasAyer['BRL'] = filaAyer[4];
      tasasAyer['ARS'] = filaAyer[5];
    }

    // ── Actualizar hoja Tasas-Hoy
    hojaHoy.clearContents();
    hojaHoy.getRange('A1:D1')
      .setValues([['Moneda', 'Tasa Actual', 'Tasa Ayer', 'Variación %']])
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    CONFIG.MONEDAS.forEach((moneda, i) => {
      const tasaHoy = tasas[moneda];
      const tasaAyer = tasasAyer[moneda] || tasaHoy;
      const variacion = tasaAyer ? (((tasaHoy - tasaAyer) / tasaAyer) * 100).toFixed(2) : '0.00';
      const signo = variacion > 0 ? '▲' : variacion < 0 ? '▼' : '─';
      const fila = i + 2;

      hojaHoy.getRange(fila, 1).setValue(`USD → ${moneda}`);
      hojaHoy.getRange(fila, 2).setValue(tasaHoy);
      hojaHoy.getRange(fila, 3).setValue(tasaAyer);
      hojaHoy.getRange(fila, 4).setValue(`${signo} ${variacion}%`);

      // Color según variación
      const color = variacion > 0 ? '#d4edda' : variacion < 0 ? '#f8d7da' : '#ffffff';
      hojaHoy.getRange(fila, 1, 1, 4).setBackground(color);
    });

    hojaHoy.autoResizeColumns(1, 4);

    // ── Agregar fila al historial
    hojaHistorial.appendRow([
      fecha,
      tasas['COP'],
      tasas['EUR'],
      tasas['MXN'],
      tasas['BRL'],
      tasas['ARS']
    ]);

    SpreadsheetApp.getUi().alert(`✅ Tasas actualizadas correctamente.\nFecha: ${fecha} ${hora}`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error al obtener tasas: ' + e.message);
    Logger.log('Error: ' + e.message);
  }
}


// ══════════════════════════════════════════════
//  ACTUALIZAR DASHBOARD
// ══════════════════════════════════════════════
function actualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHoy = ss.getSheetByName('Tasas-Hoy');
  const hojaHistorial = ss.getSheetByName('Historial');
  const hojaDashboard = ss.getSheetByName('Dashboard');

  hojaDashboard.clearContents();

  // Título
  hojaDashboard.getRange('A1').setValue('💱 MONITOR DE DIVISAS — DASHBOARD');
  hojaDashboard.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  hojaDashboard.getRange('A2').setValue(`Última actualización: ${fecha}`);
  hojaDashboard.getRange('A2').setFontColor('#666666').setFontStyle('italic');

  // Copiar tasas de hoy al dashboard
  hojaDashboard.getRange('A4').setValue('TASAS DEL DÍA');
  hojaDashboard.getRange('A4:D4').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  const datosHoy = hojaHoy.getRange(2, 1, CONFIG.MONEDAS.length, 4).getValues();
  hojaDashboard.getRange(5, 1, datosHoy.length, 4).setValues(datosHoy);

  // Encabezados de historial
  hojaDashboard.getRange('A11').setValue('HISTORIAL DE ÚLTIMOS 7 DÍAS');
  hojaDashboard.getRange('A11:F11').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  const encabezados = [['Fecha', 'COP', 'EUR', 'MXN', 'BRL', 'ARS']];
  hojaDashboard.getRange('A12:F12').setValues(encabezados).setFontWeight('bold').setBackground('#e8f0fe');

  // Últimas 7 filas del historial
  const ultimaFila = hojaHistorial.getLastRow();
  const desdeFilaHist = Math.max(2, ultimaFila - 6);
  const historialDatos = hojaHistorial.getRange(desdeFilaHist, 1, ultimaFila - desdeFilaHist + 1, 6).getValues();
  hojaDashboard.getRange(13, 1, historialDatos.length, 6).setValues(historialDatos);

  hojaDashboard.autoResizeColumns(1, 6);
  SpreadsheetApp.getUi().alert('✅ Dashboard actualizado.');
}


// ══════════════════════════════════════════════
//  GENERAR Y ENVIAR REPORTE PDF
// ══════════════════════════════════════════════
function generarYEnviarReporte() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHoy = ss.getSheetByName('Tasas-Hoy');

  try {
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

    // Leer tasas actuales
    const tasas = {};
    CONFIG.MONEDAS.forEach((moneda, i) => {
      tasas[moneda] = hojaHoy.getRange(i + 2, 2).getValue();
    });

    // Leer variaciones
    const variaciones = {};
    CONFIG.MONEDAS.forEach((moneda, i) => {
      variaciones[moneda] = hojaHoy.getRange(i + 2, 4).getValue();
    });

    // Copiar plantilla
    const carpetaTemp = DriveApp.getFolderById(CONFIG.ID_CARPETA_TEMPORAL);
    const carpetaReportes = DriveApp.getFolderById(CONFIG.ID_CARPETA_REPORTES);
    const plantilla = DriveApp.getFileById(CONFIG.ID_PLANTILLA_REPORTE);
    const copia = plantilla.makeCopy(`Reporte_${fecha}`, carpetaTemp);

    // Reemplazar placeholders
    const doc = DocumentApp.openById(copia.getId());
    const cuerpo = doc.getBody();

    cuerpo.replaceText('{{fecha}}', fecha);
    cuerpo.replaceText('{{hora}}', hora);
    cuerpo.replaceText('{{moneda_base}}', 'USD (Dólar Estadounidense)');
    cuerpo.replaceText('{{usd_cop}}', String(tasas['COP']));
    cuerpo.replaceText('{{usd_eur}}', String(tasas['EUR']));
    cuerpo.replaceText('{{usd_mxn}}', String(tasas['MXN']));
    cuerpo.replaceText('{{usd_brl}}', String(tasas['BRL']));
    cuerpo.replaceText('{{usd_ars}}', String(tasas['ARS']));

    cuerpo.replaceText('{{variacion_cop}}', String(variaciones['COP'] || '─'));
    cuerpo.replaceText('{{variacion_eur}}', String(variaciones['EUR'] || '─'));
    cuerpo.replaceText('{{variacion_mxn}}', String(variaciones['MXN'] || '─'));
    cuerpo.replaceText('{{variacion_brl}}', String(variaciones['BRL'] || '─'));
    cuerpo.replaceText('{{variacion_ars}}', String(variaciones['ARS'] || '─'));

    doc.saveAndClose();

    // Convertir a PDF y guardar
    const blobPDF = copia.getAs(MimeType.PDF);
    carpetaReportes.createFile(blobPDF).setName(`Reporte-Divisas_${fecha}.pdf`);

    // Enviar por correo
    GmailApp.sendEmail(
      CONFIG.EMAIL_DESTINATARIO,
      `Reporte de Tasas de Cambio - ${fecha}`,
      `Hola,\n\nAdjunto encontrarás el reporte de tasas de cambio del día ${fecha}.\n\nResumen:\n• USD/COP: ${tasas['COP']}\n• USD/EUR: ${tasas['EUR']}\n• USD/MXN: ${tasas['MXN']}\n• USD/BRL: ${tasas['BRL']}\n• USD/ARS: ${tasas['ARS']}\n\nReporte generado automáticamente.\nSistema Monitor de Divisas`,
      { attachments: [blobPDF], name: 'Monitor de Divisas' }
    );

    // Borrar copia temporal
    copia.setTrashed(true);

    SpreadsheetApp.getUi().alert(`✅ Reporte generado y enviado a ${CONFIG.EMAIL_DESTINATARIO}`);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error al generar reporte: ' + e.message);
    Logger.log('Error: ' + e.message);
  }
}


// ══════════════════════════════════════════════
//  FUNCIÓN AUTOMÁTICA DIARIA (para el trigger)
// ══════════════════════════════════════════════
function ejecutarDiario() {
  actualizarTasas();
  Utilities.sleep(2000);
  actualizarDashboard();
  Utilities.sleep(2000);
  generarYEnviarReporte();
}


// ══════════════════════════════════════════════
//  LIMPIAR TASAS DE HOY
// ══════════════════════════════════════════════
function limpiarTasasHoy() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasas-Hoy');
  hoja.clearContents();
  hoja.getRange('A1:D1')
    .setValues([['Moneda', 'Tasa Actual', 'Tasa Ayer', 'Variación %']])
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  SpreadsheetApp.getUi().alert('✅ Tasas limpiadas.');
}
