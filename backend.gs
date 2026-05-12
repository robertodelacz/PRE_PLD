// ══════════════════════════════════════════════════════════════
//  GENERADOR PRE-AVISO PLD — Backend.gs
//  Financiera Cualli · Contraloria Operativa
//  Genera un Spreadsheet dinamico con secciones PF/PM,
//  N vendedores, N compradores, N pagos, y protecciones y VALIDACIONES.
// ══════════════════════════════════════════════════════════════

var CONFIG = {
  CARPETA_DESTINO_ID: '1pMaxUFbXE29CJSaKRp4l90MST7cAF3eZ',

  CORREOS_COMPARTIR: [
    'robertodelacruz@cualli.mx'
  ],

  // ── Paleta Institucional Actualizada ──
  HDR_BG:        '#364494', // Azul institucional
  HDR_TEXT:      '#ffffff', // Blanco
  HDR_ACCENT:    '#35B7C5', // Cyan institucional

  LADO1_BG:      '#35B7C5', // Cyan institucional
  LADO1_TEXT:    '#ffffff', // Blanco
  LADO2_BG:      '#364494', // Azul institucional
  LADO2_TEXT:    '#ffffff', // Blanco

  SECCION_BG:    '#B2B2B2', // Gris institucional
  SECCION_TEXT:  '#ffffff', // Blanco

  LABEL_BG:      '#F4F4F8', // Gris muy claro para fondo de etiquetas
  LABEL_TEXT:    '#364494', // Azul institucional
  INPUT_BG:      '#ffffff', // Blanco
  INPUT_BORDER:  '#B2B2B2', // Gris institucional

  RESUMEN_BG:    '#F0F8F9', // Cyan ultra claro
  RESUMEN_BORDER:'#35B7C5', // Cyan institucional
  RESUMEN_TEXT:  '#364494', // Azul institucional

  PIE_BG:        '#364494', // Azul institucional
  PIE_TEXT:      '#35B7C5', // Cyan institucional
  SPACER_BG:     '#ffffff',

  ANCHOS: [80, 170, 155, 170, 155, 130, 150],

  // ── Colores seccion Documentos / Links ──
  DOC_BG:        '#F4F6F9',
  DOC_BORDER:    '#35B7C5', // Cyan institucional
  DOC_TEXT:      '#364494', // Azul institucional
  DOC_CHECK:     '#35B7C5', // Cyan institucional
  LINK_BG:       '#ffffff',
  LINK_BORDER:   '#35B7C5', // Cyan institucional
  LINK_TEXT:     '#35B7C5', // Cyan institucional
  DARK_LINK:     '#364494', // Azul institucional

  PAISES: 'MEXICO,ESTADOS UNIDOS,CANADA,ESPAÑA,COLOMBIA,ARGENTINA,BRASIL,CHILE,PERU,ALEMANIA,FRANCIA,ITALIA,REINO UNIDO,CHINA,JAPON,OTRO',
  ESTADOS: 'AGUASCALIENTES,BAJA CALIFORNIA,BAJA CALIFORNIA SUR,CAMPECHE,CHIAPAS,CHIHUAHUA,CIUDAD DE MEXICO,COAHUILA,COLIMA,DURANGO,GUANAJUATO,GUERRERO,HIDALGO,JALISCO,MEXICO,MICHOACAN,MORELOS,NAYARIT,NUEVO LEON,OAXACA,PUEBLA,QUERETARO,QUINTANA ROO,SAN LUIS POTOSI,SINALOA,SONORA,TABASCO,TAMAULIPAS,TLAXCALA,VERACRUZ,YUCATAN,ZACATECAS',
  TIPO_INMUEBLE: 'Casa / Casa en condominio,Departamento,Edificio habitacional,Edificio comercial,Edificio oficinas,Local comercial independiente,Local en centro comercial,Oficina,Bodega comercial,Bodega industrial,Nave Industrial,Terreno urbano habitacional,Terreno no urbano habitacional,Terreno urbano comercial o industrial,Terreno no urbano comercial o industrial,Terreno ejidal,Rancho/Hacienda/Quinta,Huerta,Otro',
  MONEDAS: 'Peso mexicano,Dolar estadounidense,Euro',
  FORMA_PAGO: 'Contado,Diferido o en parcialidades,Dacion en pago,Prestamo o credito,Permuta',
  INSTRUMENTO_PAGO: 'Efectivo,Tarjeta de Credito,Tarjeta de Debito,Tarjeta de Prepago,Cheque Nominativo,Cheque de Caja,Cheques de Viajero,Transferencia Interbancaria,Transferencia Misma Institucion,Transferencia Internacional,Orden de Pago,Giro,Oro o Platino Amonedados,Plata Amonedada,Metales Preciosos,Activos Virtuales,Otros',
  FIGURA_CLIENTE: 'Comprador,Vendedor'
};

function obtenerURLDelLogoPublico() {
  const FILE_ID = '1FYCSTRo_TRULxHRHfocIdUCD6Vwp1gB2';
  return 'https://drive.google.com/thumbnail?id=' + FILE_ID + '&sz=w1000';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
    .setTitle('Generador Pre-Aviso PLD')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function generarPreAvisoPLD(params) {
  try {
    var fechaHoy = Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd-MM-yyyy');
    var nombreArchivo = 'Pre_PLD_' + params.operacion.substring(0, 30) + '_' + fechaHoy;
    
    var ss = SpreadsheetApp.create(nombreArchivo);
    
    // ── NUEVO: FORZAR REGIÓN Y ZONA HORARIA A MÉXICO ──
    // Esto estandariza las fórmulas (usa comas) y la moneda (pesos) para todos los usuarios, evitando errores.
    ss.setSpreadsheetLocale('es_MX');
    ss.setSpreadsheetTimeZone('America/Mexico_City');
    // ──────────────────────────────────────────────────

    var sheet = ss.getActiveSheet();
    sheet.setName('Pre Aviso PLD');

    var archivo = DriveApp.getFileById(ss.getId());
    
    // ── DAR ACCESO INMEDIATO AL USUARIO CREADOR ──
    try {
      var userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        archivo.addEditor(userEmail);
      }
    } catch (e) {
      Logger.log('No se pudo obtener el correo del usuario: ' + e.message);
    }
    // ─────────────────────────────────────────────

    // Mover a la carpeta de destino
    try {
      var carpeta = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO_ID);
      carpeta.addFile(archivo);
      DriveApp.getRootFolder().removeFile(archivo);
    } catch (e) {
      Logger.log('Carpeta no encontrada, archivo en raiz: ' + e.message);
    }

    for (var i = 0; i < CONFIG.ANCHOS.length; i++) {
      sheet.setColumnWidth(i + 1, CONFIG.ANCHOS[i]);
    }
    sheet.setHiddenGridlines(true);

    var fila = 1;
    var celdasEditables = [];
    var celdasDropdown = [];
    var celdasRegex = []; 
    var celdasFechas = []; 
    var celdasMontos = []; 

    // ─── ENCABEZADO ───
    fila = escribirEncabezado(sheet, fila, params);

    // ─── VENDEDORES ───
    for (var v = 1; v <= params.vendedorCount; v++) {
      fila = escribirBarraLado(sheet, fila, 'LADO 1', 'VENDEDOR ' + v + ' DE ' + params.vendedorCount, CONFIG.LADO1_BG, CONFIG.LADO1_TEXT);
      if (params.vendedorRegimen === 'PF') {
        fila = escribirBloquePersonaFisica(sheet, fila, celdasEditables, celdasDropdown, celdasRegex, celdasFechas, celdasMontos);
      } else {
        fila = escribirBloquePersonaMoral(sheet, fila, celdasEditables, celdasDropdown, celdasRegex, celdasFechas, celdasMontos);
      }
      fila = escribirBarraSeccion(sheet, fila, 'Domicilio Nacional');
      fila = escribirBloqueDomicilio(sheet, fila, celdasEditables, celdasDropdown);
      fila = escribirBarraSeccion(sheet, fila, 'Datos de Contacto');
      fila = escribirBloqueDatosContacto(sheet, fila, celdasEditables, celdasDropdown);
    }

    // ─── DETALLES OPERACIÓN ───
    fila = escribirBarraLado(sheet, fila, 'LADO 1', 'DETALLES DE LA OPERACION', CONFIG.LADO1_BG, CONFIG.LADO1_TEXT);
    fila = escribirBloqueDetallesOp(sheet, fila, celdasEditables, celdasDropdown, celdasFechas);

    // ─── COMPRADORES ───
    for (var c = 1; c <= params.compradorCount; c++) {
      fila = escribirBarraLado(sheet, fila, 'LADO 2', 'COMPRADOR ' + c + ' DE ' + params.compradorCount, CONFIG.LADO2_BG, CONFIG.LADO2_TEXT);
      if (params.compradorRegimen === 'PF') {
        fila = escribirBloquePersonaFisica(sheet, fila, celdasEditables, celdasDropdown, celdasRegex, celdasFechas, celdasMontos);
      } else {
        fila = escribirBloquePersonaMoral(sheet, fila, celdasEditables, celdasDropdown, celdasRegex, celdasFechas, celdasMontos);
      }
    }

    // ─── INMUEBLE ───
    fila = escribirBarraLado(sheet, fila, 'LADO 1', 'CARACTERISTICAS DEL INMUEBLE', CONFIG.LADO1_BG, CONFIG.LADO1_TEXT);
    fila = escribirBloqueInmueble(sheet, fila, celdasEditables, celdasDropdown, celdasMontos);

    // ─── ESCRITURACIÓN ───
    fila = escribirBarraLado(sheet, fila, 'LADO 2', 'ESCRITURACION', CONFIG.LADO2_BG, CONFIG.LADO2_TEXT);
    fila = escribirBloqueEscrituracion(sheet, fila, celdasEditables, celdasDropdown, celdasFechas, celdasMontos);

    // ─── PAGOS ───
    fila = escribirBarraLado(sheet, fila, 'LADO 2', 'LIQUIDACION  —  ' + params.pagosCount + ' PAGO(S)', CONFIG.LADO2_BG, CONFIG.LADO2_TEXT);
    for (var p = 1; p <= params.pagosCount; p++) {
      fila = escribirBarraSeccion(sheet, fila, 'PAGO ' + p + ' DE ' + params.pagosCount);
      fila = escribirBloquePago(sheet, fila, celdasEditables, celdasDropdown, celdasFechas, celdasMontos);
    }

    // ─── PIE ───
    fila = escribirSpacer(sheet, fila);
    sheet.getRange(fila, 1, 1, 7).merge()
      .setValue('Gracias por tu colaboracion!')
      .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontColor(CONFIG.PIE_TEXT).setBackground(CONFIG.PIE_BG);
    sheet.setRowHeight(fila, 44);
    fila++;
    sheet.getRange(fila, 1, 1, 7).merge().setBackground(CONFIG.HDR_ACCENT);
    sheet.setRowHeight(fila, 4);

    aplicarDropdowns(sheet, celdasDropdown);
    aplicarProtecciones(sheet, fila, celdasEditables); // Ahora funciona con expulsión de editores
    aplicarValidacionesRegex(sheet, celdasRegex);      // Ahora usando [0-9]
    aplicarValidacionesFechas(sheet, celdasFechas);
    aplicarValidacionMontos(sheet, celdasMontos);
    aplicarSemaforo(sheet, celdasEditables); 
    
    compartirArchivo(archivo);
    
    sheet.setFrozenRows(2);

    return { success: true, url: ss.getUrl(), id: ss.getId() };
  } catch (e) {
    Logger.log('Error: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
//  ENCABEZADO
// ══════════════════════════════════════════════════════════════

function escribirEncabezado(sheet, fila, params) {
  sheet.getRange(fila, 1, 1, 7).merge().setBackground(CONFIG.HDR_ACCENT);
  sheet.setRowHeight(fila, 4);
  fila++;

  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('PRE AVISO PLD  /  Entregable max 48 hrs PREVIAS a firma de Escritura')
    .setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
    .setFontColor(CONFIG.HDR_TEXT).setBackground(CONFIG.HDR_BG)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(fila, 48);
  fila++;

  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('Recordemos que el momento de la escrituracion es el ULTIMO que tenemos para recabar informacion faltante.')
    .setFontFamily('Arial').setFontSize(9).setFontStyle('italic')
    .setFontColor(CONFIG.HDR_ACCENT).setBackground(CONFIG.HDR_BG)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(fila, 30);
  fila++;

  sheet.getRange(fila, 1, 1, 7).merge().setBackground(CONFIG.HDR_ACCENT);
  sheet.setRowHeight(fila, 3);
  fila++;

  fila = escribirSpacer(sheet, fila);

  fila = escribirCampoHeader(sheet, fila, 'Operacion:', params.operacion);

  sheet.getRange(fila, 1, 1, 2).merge()
    .setValue('Asociados Responsables:')
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor(CONFIG.LABEL_TEXT).setBackground(CONFIG.LABEL_BG).setVerticalAlignment('middle');
  sheet.getRange(fila, 3, 1, 2).merge()
    .setValue('Lado 1:  ' + (params.asociadoLado1 || ''))
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(fila, 5, 1, 3).merge()
    .setValue('Lado 2:  ' + (params.asociadoLado2 || ''))
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(fila, 28);
  fila++;

  sheet.getRange(fila, 1, 1, 2).merge()
    .setValue('Liga a Expediente Digital:')
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor(CONFIG.LABEL_TEXT).setBackground(CONFIG.LABEL_BG).setVerticalAlignment('middle');
  sheet.getRange(fila, 3, 1, 5).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(fila, 28);
  fila++;

  fila = escribirSpacer(sheet, fila);

  var rl = { 'PF': 'Persona Fisica', 'PM': 'Persona Moral' };
  var resumen = '▸ ' + params.vendedorCount + ' Vendedor(es) ' + rl[params.vendedorRegimen]
    + '    ▸ ' + params.compradorCount + ' Comprador(es) ' + rl[params.compradorRegimen]
    + '    ▸ ' + params.pagosCount + ' Pago(s)';

  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue(resumen)
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(CONFIG.RESUMEN_TEXT).setBackground(CONFIG.RESUMEN_BG)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(fila, 34);
  sheet.getRange(fila, 1).setBorder(true, true, true, false, false, false, CONFIG.RESUMEN_BORDER, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(fila, 7).setBorder(true, false, true, true, false, false, CONFIG.RESUMEN_BORDER, SpreadsheetApp.BorderStyle.SOLID_THICK);
  fila++;

  fila = escribirSpacer(sheet, fila);
  fila = escribirSeccionDocumentos(sheet, fila);
  fila = escribirSpacer(sheet, fila);
  fila = escribirSeccionEnlaces(sheet, fila);
  fila = escribirSpacer(sheet, fila);

  return fila;
}

function escribirCampoHeader(sheet, fila, label, valor) {
  sheet.getRange(fila, 1, 1, 2).merge()
    .setValue(label)
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor(CONFIG.LABEL_TEXT).setBackground(CONFIG.LABEL_BG).setVerticalAlignment('middle');
  sheet.getRange(fila, 3, 1, 5).merge()
    .setValue(valor || '')
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(fila, 28);
  return fila + 1;
}

// ══════════════════════════════════════════════════════════════
//  DOCUMENTOS REQUERIDOS + ENLACES
// ══════════════════════════════════════════════════════════════

function escribirSeccionDocumentos(sheet, fila) {
  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('📋  DOCUMENTOS REQUERIDOS')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(CONFIG.DOC_TEXT).setBackground(CONFIG.DOC_BG)
    .setVerticalAlignment('middle');
  sheet.getRange(fila, 1).setBorder(true, true, false, false, false, false, CONFIG.DOC_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(fila, 7).setBorder(true, false, false, true, false, false, CONFIG.DOC_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(fila, 28);
  fila++;

  var docs = [
    ['✓  La nueva escritura de compraventa',       '✓  Comprobantes de pago por monto total de la operacion'],
    ['✓  Identificacion oficial vigente',            '✓  Formatos PLD firmados'],
    ['✓  CSF Constancia de Situacion Fiscal',       '']
  ];

  for (var i = 0; i < docs.length; i++) {
    sheet.getRange(fila, 1, 1, 3).merge()
      .setValue(docs[i][0])
      .setFontFamily('Arial').setFontSize(9)
      .setFontColor(CONFIG.DOC_TEXT).setBackground(CONFIG.DOC_BG)
      .setVerticalAlignment('middle');
    sheet.getRange(fila, 4, 1, 4).merge()
      .setValue(docs[i][1])
      .setFontFamily('Arial').setFontSize(9)
      .setFontColor(CONFIG.DOC_TEXT).setBackground(CONFIG.DOC_BG)
      .setVerticalAlignment('middle');
    
    sheet.getRange(fila, 1).setBorder(false, true, false, false, false, false, CONFIG.DOC_BORDER, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(fila, 7).setBorder(false, false, false, true, false, false, CONFIG.DOC_BORDER, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(fila, 22);
    fila++;
  }

  sheet.getRange(fila - 1, 1, 1, 7).setBorder(false, true, true, true, false, false, CONFIG.DOC_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  return fila;
}

function escribirSeccionEnlaces(sheet, fila) {
  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue('📄  Catalogo de Actividades Economicas')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(CONFIG.HDR_TEXT).setBackground(CONFIG.DARK_LINK)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, CONFIG.DARK_LINK, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(fila, 4, 1, 4).merge()
    .setFontFamily('Arial').setFontSize(9)
    .setFontColor(CONFIG.HDR_ACCENT).setBackground(CONFIG.DARK_LINK)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, CONFIG.DARK_LINK, SpreadsheetApp.BorderStyle.SOLID);
  
  var estiloLink = SpreadsheetApp.newTextStyle()
    .setForegroundColor('#ffffff') 
    .setUnderline(true)            
    .build();
    
  var rt1 = SpreadsheetApp.newRichTextValue()
    .setText('▸ Abrir catalogo en Google Sheets')
    .setLinkUrl(2, 33, 'https://docs.google.com/spreadsheets/d/1GjQp9PQs8ZD09K5--YEFM6-sXT2MjSR4IaqOocQyn7s/edit')
    .setTextStyle(2, 33, estiloLink) 
    .build();
  sheet.getRange(fila, 4).setRichTextValue(rt1);
  sheet.setRowHeight(fila, 32);
  fila++;

  fila = escribirSpacer(sheet, fila);

  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue('📝  Formulario de Operaciones Firmadas')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(CONFIG.HDR_TEXT).setBackground(CONFIG.DARK_LINK)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, false, false, false, CONFIG.DARK_LINK, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(fila, 4, 1, 4).merge()
    .setFontFamily('Arial').setFontSize(9)
    .setFontColor(CONFIG.HDR_ACCENT).setBackground(CONFIG.DARK_LINK)
    .setVerticalAlignment('middle')
    .setBorder(true, false, true, true, false, false, CONFIG.DARK_LINK, SpreadsheetApp.BorderStyle.SOLID);
    
  var rt2 = SpreadsheetApp.newRichTextValue()
    .setText('▸ Abrir formulario en Google Forms')
    .setLinkUrl(2, 34, 'https://docs.google.com/forms/d/e/1FAIpQLSchuuGFfwQta-UXEt0aFUG4ZTLOAZVOJNrxm6pb09UiR5E4UA/viewform')
    .setTextStyle(2, 34, estiloLink)
    .build();
  sheet.getRange(fila, 4).setRichTextValue(rt2);
  sheet.setRowHeight(fila, 32);
  fila++;

  return fila;
}

// ══════════════════════════════════════════════════════════════
//  BARRAS
// ══════════════════════════════════════════════════════════════

function escribirBarraLado(sheet, fila, lado, texto, bg, textColor) {
  sheet.getRange(fila, 1)
    .setValue(lado)
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(textColor).setBackground(bg)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(fila, 2, 1, 6).merge()
    .setValue(texto)
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor(textColor).setBackground(bg).setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 34);
  return fila + 1;
}

function escribirBarraSeccion(sheet, fila, texto) {
  sheet.getRange(fila, 1).setBackground(CONFIG.SECCION_BG);
  sheet.getRange(fila, 2, 1, 6).merge()
    .setValue(texto)
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor(CONFIG.SECCION_TEXT).setBackground(CONFIG.SECCION_BG).setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 26);
  return fila + 1;
}

// ══════════════════════════════════════════════════════════════
//  BLOQUES
// ══════════════════════════════════════════════════════════════

function escribirBloquePersonaFisica(sheet, fila, ed, dd, cr, cf, cm) {
  fila = escribirBarraSeccion(sheet, fila, 'Persona Fisica');
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Nombres:', col: 2, inputCol: 3 },
    { label: 'Ap. Paterno:', col: 4, inputCol: 5 },
    { label: 'Ap. Materno:', col: 6, inputCol: 7 }
  ], cf, cm);
  
  var fVal = fila;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Fecha Nac (dd/mm/aaaa):', col: 2, inputCol: 3, type: 'fecha' },
    { label: 'RFC (13 Posiciones):', col: 4, inputCol: 5 },
    { label: 'CURP (18 Posiciones):', col: 6, inputCol: 7 }
  ], cf, cm);
  cr.push({ cell: sheet.getRange(fVal, 5).getA1Notation(), type: 'RFC_PF' });
  cr.push({ cell: sheet.getRange(fVal, 7).getA1Notation(), type: 'CURP' });

  setLabelCell(sheet, fila, 2, 'Pais de Nacionalidad:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.PAISES });
  setLabelCell(sheet, fila, 4, 'Actividad Economica:');
  sheet.getRange(fila, 5, 1, 3).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ed.push(sheet.getRange(fila, 5, 1, 3).getA1Notation());
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloquePersonaMoral(sheet, fila, ed, dd, cr, cf, cm) {
  fila = escribirBarraSeccion(sheet, fila, 'Persona Moral');
  var fVal1 = fila;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Denominacion / Razon Social:', col: 2, inputCol: 3 },
    { label: 'Fecha Constitucion:', col: 4, inputCol: 5, type: 'fecha' },
    { label: 'RFC (12 Posiciones):', col: 6, inputCol: 7 }
  ], cf, cm);
  cr.push({ cell: sheet.getRange(fVal1, 7).getA1Notation(), type: 'RFC_PM' });

  setLabelCell(sheet, fila, 2, 'Actividad o Giro Mercantil:');
  sheet.getRange(fila, 3, 1, 2).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ed.push(sheet.getRange(fila, 3, 1, 2).getA1Notation());
  setLabelCell(sheet, fila, 5, 'Pais de Nacionalidad:');
  setInputCell(sheet, fila, 6, ed);
  dd.push({ fila: fila, col: 6, opciones: CONFIG.PAISES });
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  
  fila = escribirBarraSeccion(sheet, fila, 'Representante o Apoderado Legal');
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Nombres:', col: 2, inputCol: 3 },
    { label: 'Ap. Paterno:', col: 4, inputCol: 5 },
    { label: 'Ap. Materno:', col: 6, inputCol: 7 }
  ], cf, cm);
  
  var fVal2 = fila;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Fecha Nacimiento (dd/mm/aaaa):', col: 2, inputCol: 3, type: 'fecha' },
    { label: 'RFC (13 Posiciones):', col: 4, inputCol: 5 },
    { label: 'CURP (18 Posiciones):', col: 6, inputCol: 7 }
  ], cf, cm);
  cr.push({ cell: sheet.getRange(fVal2, 5).getA1Notation(), type: 'RFC_PF' });
  cr.push({ cell: sheet.getRange(fVal2, 7).getA1Notation(), type: 'CURP' });

  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloqueDomicilio(sheet, fila, ed, dd) {
  setLabelCell(sheet, fila, 2, 'Entidad Federativa:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.ESTADOS });
  setLabelCell(sheet, fila, 4, 'Municipio / Alcaldía:');
  setInputCell(sheet, fila, 5, ed);
  setLabelCell(sheet, fila, 6, 'Colonia:');
  setInputCell(sheet, fila, 7, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Calle, avenida o via:', col: 2, inputCol: 3 },
    { label: 'Num. Ext:', col: 4, inputCol: 5 },
    { label: 'Num. Int:', col: 6, inputCol: 7 }
  ]);
  setLabelCell(sheet, fila, 2, 'C.P.:');
  setInputCell(sheet, fila, 3, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloqueDatosContacto(sheet, fila, ed, dd) {
  setLabelCell(sheet, fila, 2, 'Pais:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.PAISES });
  setLabelCell(sheet, fila, 4, 'Tel. Movil:');
  setInputCell(sheet, fila, 5, ed);
  setLabelCell(sheet, fila, 6, 'Correo electronico:');
  setInputCell(sheet, fila, 7, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloqueDetallesOp(sheet, fila, ed, dd, cf) {
  setLabelCell(sheet, fila, 2, 'Fecha de la Operación (dd/mm/aaaa):');
  var a1 = setInputCell(sheet, fila, 3, ed);
  cf.push(a1);
  setLabelCell(sheet, fila, 4, 'Tipo de la Operación:');
  setInputCell(sheet, fila, 5, ed);
  setLabelCell(sheet, fila, 6, 'Figura del Cliente Reportado:');
  setInputCell(sheet, fila, 7, ed);
  dd.push({ fila: fila, col: 7, opciones: CONFIG.FIGURA_CLIENTE });
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloqueInmueble(sheet, fila, ed, dd, cm) {
  setLabelCell(sheet, fila, 2, 'Tipo de inmueble:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.TIPO_INMUEBLE });
  setLabelCell(sheet, fila, 4, 'Valor Pactado: $');
  var a1 = setInputCell(sheet, fila, 5, ed);
  cm.push(a1);
  setLabelCell(sheet, fila, 6, 'Folio Real:');
  setInputCell(sheet, fila, 7, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'm2 Terreno:', col: 2, inputCol: 3 },
    { label: 'm2 Construccion:', col: 4, inputCol: 5 }
  ]);
  fila = escribirSpacer(sheet, fila);
  fila = escribirBarraSeccion(sheet, fila, 'Direccion del Inmueble');
  setLabelCell(sheet, fila, 2, 'Entidad Federativa:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.ESTADOS });
  setLabelCell(sheet, fila, 4, 'Municipio / Alcaldía:');
  setInputCell(sheet, fila, 5, ed);
  setLabelCell(sheet, fila, 6, 'Colonia:');
  setInputCell(sheet, fila, 7, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Calle:', col: 2, inputCol: 3 },
    { label: 'Num. Ext:', col: 4, inputCol: 5 },
    { label: 'Num. Int:', col: 6, inputCol: 7 }
  ]);
  setLabelCell(sheet, fila, 2, 'C.P.:');
  setInputCell(sheet, fila, 3, ed);
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloqueEscrituracion(sheet, fila, ed, dd, cf, cm) {
  setLabelCell(sheet, fila, 2, 'Num. del Instrumento Publico:');
  setInputCell(sheet, fila, 3, ed);
  setLabelCell(sheet, fila, 4, 'Fecha del Instrumento:');
  sheet.getRange(fila, 5, 1, 3).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  var a1_fecha = sheet.getRange(fila, 5, 1, 3).getA1Notation();
  ed.push(a1_fecha);
  cf.push(a1_fecha);
  
  sheet.setRowHeight(fila, 28);
  fila++;
  setLabelCell(sheet, fila, 2, 'Num. Notario:');
  setInputCell(sheet, fila, 3, ed);
  setLabelCell(sheet, fila, 4, 'Entidad Fed. del Notario:');
  setInputCell(sheet, fila, 5, ed);
  dd.push({ fila: fila, col: 5, opciones: CONFIG.ESTADOS });
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirFilaCampos(sheet, fila, ed, [
    { label: 'Valor Avaluo: $', col: 2, inputCol: 3, type: 'monto' },
    { label: 'Valor Catastral: $', col: 4, inputCol: 5, type: 'monto' }
  ], cf, cm);
  fila = escribirSpacer(sheet, fila);
  return fila;
}

function escribirBloquePago(sheet, fila, ed, dd, cf, cm) {
  setLabelCell(sheet, fila, 2, 'Fecha de Pago (dd/mm/aaaa):');
  var a1_fecha = setInputCell(sheet, fila, 3, ed);
  cf.push(a1_fecha);
  setLabelCell(sheet, fila, 4, 'Comprobante de Pago:');
  sheet.getRange(fila, 5, 1, 3).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ed.push(sheet.getRange(fila, 5, 1, 3).getA1Notation());
  sheet.setRowHeight(fila, 28);
  fila++;
  setLabelCell(sheet, fila, 2, 'Monto del pago:');
  var a1_monto = setInputCell(sheet, fila, 3, ed);
  cm.push(a1_monto);
  setLabelCell(sheet, fila, 4, 'Moneda o Divisa:');
  setInputCell(sheet, fila, 5, ed);
  dd.push({ fila: fila, col: 5, opciones: CONFIG.MONEDAS });
  sheet.setRowHeight(fila, 28);
  fila++;
  setLabelCell(sheet, fila, 2, 'Forma de Pago:');
  setInputCell(sheet, fila, 3, ed);
  dd.push({ fila: fila, col: 3, opciones: CONFIG.FORMA_PAGO });
  sheet.setRowHeight(fila, 28);
  fila++;
  setLabelCell(sheet, fila, 2, 'Instrumento de Pago:');
  sheet.getRange(fila, 3, 1, 5).merge()
    .setBackground(CONFIG.INPUT_BG).setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(CONFIG.HDR_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ed.push(sheet.getRange(fila, 3, 1, 5).getA1Notation());
  dd.push({ fila: fila, col: 3, opciones: CONFIG.INSTRUMENTO_PAGO });
  sheet.setRowHeight(fila, 28);
  fila++;
  fila = escribirSpacer(sheet, fila);
  return fila;
}

// ══════════════════════════════════════════════════════════════
//  HELPERS (Validaciones y Componentes Base)
// ══════════════════════════════════════════════════════════════

// ── Convierte Fechas Dinámicamente ──
function aplicarValidacionesFechas(sheet, celdasFechas) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false) // Rechaza texto no válido
    .setHelpText('Ingresa una fecha válida. El sistema la convertirá automáticamente a DD/MM/AAAA.')
    .build();

  for (var i = 0; i < celdasFechas.length; i++) {
    var range = sheet.getRange(celdasFechas[i]);
    range.setDataValidation(rule);
    range.setNumberFormat('dd/MM/yyyy'); // Obliga la visualización a usar diagonales
  }
}

// ── Blinda y Formatea Montos ──
function aplicarValidacionMontos(sheet, celdasMontos) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false) // Rechaza letras o signos
    .setHelpText('Solo números. El formato de moneda ($) se aplicará automáticamente.')
    .build();

  for (var i = 0; i < celdasMontos.length; i++) {
    var range = sheet.getRange(celdasMontos[i]);
    range.setDataValidation(rule);
    range.setNumberFormat('"$ "#,##0.00'); // Automatiza símbolo y comas
  }
}

// ── Semáforo Visual (Formato Condicional) ──
function aplicarSemaforo(sheet, celdasEditables) {
  if (celdasEditables.length === 0) return;
  // Convertimos el arreglo de notaciones A1 en objetos Rango de Sheets
  var ranges = celdasEditables.map(function(a1) { return sheet.getRange(a1); });
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty() // Si la celda está vacía...
    .setBackground('#FFEBEB') // Se pone de un rojo tenue para llamar la atención
    .setRanges(ranges)
    .build();
    
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// ── Validación Regex para RFC y CURP ──
function aplicarValidacionesRegex(sheet, celdasRegex) {
  for (var i = 0; i < celdasRegex.length; i++) {
    var item = celdasRegex[i];
    var range = sheet.getRange(item.cell);
    var formula = "";
    var mensaje = "";

    if (item.type === 'RFC_PF') {
      // Cambiamos \\d por [0-9]
      formula = '=OR(ISBLANK(' + item.cell + '), REGEXMATCH(UPPER(TO_TEXT(' + item.cell + ')), "^[A-ZÑ&]{4}[0-9]{6}[A-Z0-9]{3}$"))';
      mensaje = "El RFC de Persona Física requiere 13 posiciones: 4 letras, 6 números (YYMMDD) y 3 caracteres de homoclave.";
    } else if (item.type === 'RFC_PM') {
      // Cambiamos \\d por [0-9]
      formula = '=OR(ISBLANK(' + item.cell + '), REGEXMATCH(UPPER(TO_TEXT(' + item.cell + ')), "^[A-ZÑ&]{3}[0-9]{6}[A-Z0-9]{3}$"))';
      mensaje = "El RFC de Persona Moral requiere 12 posiciones: 3 letras, 6 números (YYMMDD) y 3 caracteres de homoclave.";
    } else if (item.type === 'CURP') {
      // Cambiamos \\d por [0-9]
      formula = '=OR(ISBLANK(' + item.cell + '), REGEXMATCH(UPPER(TO_TEXT(' + item.cell + ')), "^[A-Z]{4}[0-9]{6}[HM][A-Z]{5}[A-Z0-9]{2}$"))';
      mensaje = "La CURP requiere exactamente 18 posiciones alfanuméricas con la estructura oficial del gobierno.";
    }

    if(formula !== "") {
      var rule = SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied(formula)
        .setAllowInvalid(false) 
        .setHelpText(mensaje)
        .build();
      range.setDataValidation(rule);
    }
  }
}

function setLabelCell(sheet, fila, col, texto) {
  sheet.getRange(fila, col)
    .setValue(texto)
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor(CONFIG.LABEL_TEXT).setBackground(CONFIG.LABEL_BG)
    .setVerticalAlignment('middle').setWrap(true);
}

function setInputCell(sheet, fila, col, editables) {
  var cell = sheet.getRange(fila, col);
  cell.setBackground(CONFIG.INPUT_BG)
    .setBorder(true, true, true, true, false, false, CONFIG.INPUT_BORDER, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setFontFamily('Arial').setFontSize(10)
    .setFontWeight('bold').setFontColor(CONFIG.HDR_BG);
  
  var a1Notation = cell.getA1Notation();
  editables.push(a1Notation);
  return a1Notation;
}

// Modificamos esta función para que reciba y mapee los arreglos de fechas y montos
function escribirFilaCampos(sheet, fila, editables, campos, cf, cm) {
  for (var i = 0; i < campos.length; i++) {
    setLabelCell(sheet, fila, campos[i].col, campos[i].label);
    var a1 = setInputCell(sheet, fila, campos[i].inputCol, editables);
    
    if (campos[i].type === 'fecha' && cf) cf.push(a1);
    if (campos[i].type === 'monto' && cm) cm.push(a1);
  }
  sheet.setRowHeight(fila, 28);
  return fila + 1;
}

function escribirSpacer(sheet, fila) {
  sheet.getRange(fila, 1, 1, 7).merge().setBackground(CONFIG.SPACER_BG);
  sheet.setRowHeight(fila, 8);
  return fila + 1;
}

function aplicarDropdowns(sheet, dropdowns) {
  for (var i = 0; i < dropdowns.length; i++) {
    var dd = dropdowns[i];
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(dd.opciones.split(','), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(dd.fila, dd.col).setDataValidation(regla);
  }
}

function aplicarProtecciones(sheet, ultimaFila, celdasEditables) {
  var proteccion = sheet.protect().setDescription('Proteccion Pre-Aviso PLD');
  var rangos = [];
  
  for (var i = 0; i < celdasEditables.length; i++) {
    rangos.push(sheet.getRange(celdasEditables[i]));
  }
  
  proteccion.setUnprotectedRanges(rangos);
  proteccion.setWarningOnly(false); // Falso = Bloqueo estricto, no solo advertencia
  
  try {
    var editoresActuales = proteccion.getEditors();
    
    proteccion.removeEditors(editoresActuales);
    
    if (CONFIG.CORREOS_COMPARTIR && CONFIG.CORREOS_COMPARTIR.length > 0) {
      proteccion.addEditors(CONFIG.CORREOS_COMPARTIR);
    }
  } catch (e) {
    Logger.log('Error al configurar los editores de la protección: ' + e.message);
  }
}

function compartirArchivo(archivo) {
  for (var i = 0; i < CONFIG.CORREOS_COMPARTIR.length; i++) {
    try {
      archivo.addEditor(CONFIG.CORREOS_COMPARTIR[i]);
    } catch (e) {
      Logger.log('Error compartir ' + CONFIG.CORREOS_COMPARTIR[i] + ': ' + e.message);
    }
  }
}
