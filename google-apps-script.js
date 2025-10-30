function doPost(e) {
  try {
    // ID de tu Google Sheets
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    
    // Verificar si la hoja tiene cabeceras
    checkAndCreateHeaders(sheet);
    
    // Parsear los datos según el formato recibido
    let data;
    
    // Si vienen como JSON en el body
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseError) {
        // Si no es JSON, intentar con parámetros
        data = e.parameter;
      }
    } else {
      // Si vienen como parámetros de formulario
      data = e.parameter;
    }
    
    // Log para debugging
    console.log('Datos recibidos:', JSON.stringify(data));
    
    // Manejar acciones específicas para emisores
    if (data.action === 'saveEmisor') {
      return handleSaveEmisor(data.data);
    }
    
    // Manejar reordenamiento de filas
    if (data.action === 'reorderRows') {
      return handleReorderRows(data);
    }
    
    // ✅ EJECUTAR INSTRUCCIONES DEL HTML: Si el HTML indica que hay filas que eliminar, ejecutar
    if (data.deleteRows && Array.isArray(data.deleteRows) && data.deleteRows.length > 0) {
      console.log('📋 HTML solicita eliminar ' + data.deleteRows.length + ' filas antiguas...');
      // Eliminar de mayor a menor para no desplazar índices
      for (let i = data.deleteRows.length - 1; i >= 0; i--) {
        sheet.deleteRow(data.deleteRows[i]);
      }
      console.log('✅ Filas eliminadas correctamente');
    }
    
    // Si hay items detallados, crear una fila por cada item
    if (data.itemsDetail && Array.isArray(data.itemsDetail)) {
      // Crear una fila por cada línea de detalle con el nuevo formato
      data.itemsDetail.forEach((item, index) => {
        const row = [
          data.fecha || '',           // Fecha
          data.serie || '',           // Serie
          data.numero || '',          // Número
          data.nif || '',             // NIF
          data.cliente || '',         // Cliente
          data.direccion || '',       // Dirección
          data.cp || '',              // CP
          data.ciudad || '',          // Ciudad
          data.provincia || '',       // Provincia
          data.email || '',           // Email
          data.descripcion || '',     // Descripción
          data.textoLibre || '',      // Texto Libre
          item.descripcion || '',     // Detalle
          item.cantidad || '',        // Cantidad
          item.precio || '',          // Precio
          item.subtotal || '',        // Base Imponible
          item.iva || '',             // IVA
          item.total || ''            // Total
        ];
        sheet.appendRow(row);
      });
    } else {
      // Modo clásico: una sola fila (cuando no hay detalle de items)
      const row = [
        data.fecha || '',           // Fecha
        data.serie || '',           // Serie
        data.numero || '',          // Número
        data.nif || '',             // NIF
        data.cliente || '',         // Cliente
        data.direccion || '',       // Dirección
        data.cp || '',              // CP
        data.ciudad || '',          // Ciudad
        data.provincia || '',       // Provincia
        data.email || '',           // Email
        data.descripcion || '',     // Descripción
        data.textoLibre || '',      // Texto Libre
        data.items || '',           // Detalle (todos concatenados)
        '',                         // Cantidad
        '',                         // Precio
        data.base || '',            // Base Imponible
        data.iva || '',             // IVA
        data.total || ''            // Total
      ];
      sheet.appendRow(row);
    }
    
    // Devolver respuesta exitosa
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: 'Factura guardada correctamente',
        row: sheet.getLastRow()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    // En caso de error, devolver información del error
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        error: error.toString(),
        message: 'Error al guardar la factura',
        details: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Función para verificar y crear cabeceras si no existen
function checkAndCreateHeaders(sheet) {
  // Verificar si la primera fila está vacía o no tiene cabeceras
  const firstRow = sheet.getRange(1, 1, 1, 18).getValues()[0];
  const hasHeaders = firstRow.some(cell => cell !== '');
  
  if (!hasHeaders || firstRow[0] !== 'Fecha') {
    // Crear cabeceras según el nuevo formato
    const headers = [
      'Fecha',
      'Serie',
      'Número',
      'NIF',
      'Cliente',
      'Dirección',
      'CP',
      'Ciudad',
      'Provincia',
      'Email',
      'Descripción',
      'Texto Libre',
      'Detalle',
      'Cantidad',
      'Precio',
      'Base Imponible',
      'IVA',
      'Total'
    ];
    
    // Si ya hay datos, insertar una fila al principio
    if (hasHeaders) {
      sheet.insertRowBefore(1);
    }
    
    // Establecer las cabeceras
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Aplicar formato a las cabeceras
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#3498db');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  }
}

// Función para manejar solicitudes GET
function doGet(e) {
  const action = e.parameter.action;
  
  try {
    switch(action) {
      case 'read':
        return handleRead();
      case 'update':
        return handleUpdate(e.parameter);
      case 'delete':
        return handleDelete(e.parameter);
      case 'getEmisores':
        return handleGetEmisores();
      case 'createEmisoresSheet':
        return handleCreateEmisoresSheet();
      case 'deleteEmisor':
        return handleDeleteEmisor(e.parameter);
      default:
        return ContentService
          .createTextOutput(JSON.stringify({
            status: 'active',
            message: 'Web App funcionando correctamente'
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Leer todas las facturas
function handleRead() {
  const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Obtener encabezados
  const headers = data[0];
  const rows = [];
  
  // Convertir a objetos
  for (let i = 1; i < data.length; i++) {
    const row = {};
    row['rowNumber'] = i + 1; // Número de fila real en Sheets
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      result: 'success',
      data: rows,
      headers: headers
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Actualizar una fila
function handleUpdate(params) {
  const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const rowNumber = parseInt(params.row);
  const columnNumber = parseInt(params.col);
  const value = params.value;
  
  // Actualizar celda
  sheet.getRange(rowNumber, columnNumber).setValue(value);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      result: 'success',
      message: 'Celda actualizada'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Eliminar una fila
function handleDelete(params) {
  const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  
  // Si hay un parámetro 'rows', es eliminación masiva
  if (params.rows) {
    return handleBulkDelete(params);
  }
  
  // Eliminación individual
  const rowNumber = parseInt(params.row);
  sheet.deleteRow(rowNumber);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      result: 'success',
      message: 'Fila eliminada'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Eliminar múltiples filas de forma eficiente
function handleBulkDelete(params) {
  const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  
  try {
    // Parsear las filas a eliminar
    const rowsToDelete = JSON.parse(params.rows);
    
    if (!Array.isArray(rowsToDelete) || rowsToDelete.length === 0) {
      throw new Error('No se proporcionaron filas válidas para eliminar');
    }
    
    // Ordenar de mayor a menor para eliminar desde abajo hacia arriba
    const sortedRows = rowsToDelete.map(r => parseInt(r)).sort((a, b) => b - a);
    
    // Agrupar filas consecutivas para usar deleteRows() de forma eficiente
    const groups = [];
    let currentGroup = { start: sortedRows[0], count: 1 };
    
    for (let i = 1; i < sortedRows.length; i++) {
      if (sortedRows[i] === sortedRows[i-1] - 1) {
        // Fila consecutiva, expandir el grupo actual
        currentGroup.start = sortedRows[i];
        currentGroup.count++;
      } else {
        // Nueva secuencia, guardar grupo actual y crear uno nuevo
        groups.push(currentGroup);
        currentGroup = { start: sortedRows[i], count: 1 };
      }
    }
    // No olvidar el último grupo
    groups.push(currentGroup);
    
    // Eliminar cada grupo usando deleteRows() para máxima eficiencia
    let deletedCount = 0;
    for (const group of groups) {
      sheet.deleteRows(group.start, group.count);
      deletedCount += group.count;
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: `${deletedCount} filas eliminadas correctamente`,
        deletedCount: deletedCount
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        error: error.toString(),
        message: 'Error en eliminación masiva'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==============================================
// HANDLERS PARA EMISORES (GET/POST)
// ==============================================

// Handler para obtener emisores
function handleGetEmisores() {
  const result = getEmisores();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handler para crear hoja de emisores
function handleCreateEmisoresSheet() {
  const result = createEmisoresSheet();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handler para eliminar emisor
function handleDeleteEmisor(params) {
  const result = deleteEmisor(params.id);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handler para guardar emisor
function handleSaveEmisor(emisorData) {
  console.log('handleSaveEmisor recibió:', JSON.stringify(emisorData));
  const result = saveEmisor(emisorData);
  console.log('saveEmisor resultado:', JSON.stringify(result));
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================
// GESTIÓN DE EMISORES
// ==============================================

// Crear hoja de emisores si no existe
function createEmisoresSheet() {
  try {
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Verificar si ya existe la hoja de emisores
    let emisoresSheet = spreadsheet.getSheetByName('Emisores');
    
    if (!emisoresSheet) {
      // Crear nueva hoja de emisores
      emisoresSheet = spreadsheet.insertSheet('Emisores');
      
      // Establecer cabeceras
      const headers = [
        'ID', 'Nombre', 'NIF', 'Teléfono', 'Email',
        'Dirección', 'Ciudad', 'CP', 'Provincia', 'Fecha Creación'
      ];
      
      emisoresSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Formatear cabeceras
      const headerRange = emisoresSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#34495e');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // Congelar primera fila
      emisoresSheet.setFrozenRows(1);
      
      // Ajustar ancho de columnas
      emisoresSheet.autoResizeColumns(1, headers.length);
      
      console.log('Hoja de emisores creada correctamente');
    }
    
    return {
      success: true,
      sheetName: 'Emisores',
      message: 'Hoja de emisores lista'
    };
    
  } catch (error) {
    console.error('Error al crear hoja de emisores:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Obtener todos los emisores
function getEmisores() {
  try {
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let emisoresSheet = spreadsheet.getSheetByName('Emisores');
    
    if (!emisoresSheet) {
      // La hoja no existe
      return {
        success: false,
        error: 'La hoja de emisores no existe'
      };
    }
    
    const lastRow = emisoresSheet.getLastRow();
    
    if (lastRow <= 1) {
      // Solo hay cabeceras o la hoja está vacía
      return {
        success: true,
        data: []
      };
    }
    
    // Obtener todos los datos
    const range = emisoresSheet.getRange(2, 1, lastRow - 1, 10);
    const values = range.getValues();
    
    const emisores = values.map(row => ({
      id: row[0] || '',
      nombre: row[1] || '',
      nif: row[2] || '',
      telefono: row[3] || '',
      email: row[4] || '',
      direccion: row[5] || '',
      ciudad: row[6] || '',
      cp: row[7] || '',
      provincia: row[8] || '',
      fechaCreacion: row[9] || ''
    })).filter(emisor => emisor.id); // Filtrar filas vacías
    
    return {
      success: true,
      data: emisores
    };
    
  } catch (error) {
    console.error('Error al obtener emisores:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Guardar o actualizar emisor
function saveEmisor(emisorData) {
  try {
    console.log('saveEmisor iniciando con datos:', JSON.stringify(emisorData));
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let emisoresSheet = spreadsheet.getSheetByName('Emisores');
    
    if (!emisoresSheet) {
      // Crear la hoja si no existe
      const createResult = createEmisoresSheet();
      if (!createResult.success) {
        return createResult;
      }
      emisoresSheet = spreadsheet.getSheetByName('Emisores');
    }
    
    const lastRow = emisoresSheet.getLastRow();
    
    // Buscar si el emisor ya existe (por ID)
    let existingRow = -1;
    if (lastRow > 1) {
      const ids = emisoresSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      existingRow = ids.findIndex(id => id == emisorData.id); // Usar == para comparación flexible
    }
    
    const rowData = [
      emisorData.id,
      emisorData.nombre,
      emisorData.nif,
      emisorData.telefono,
      emisorData.email,
      emisorData.direccion,
      emisorData.ciudad,
      emisorData.cp,
      emisorData.provincia,
      existingRow === -1 ? new Date().toISOString() : '' // Solo fecha de creación para nuevos
    ];
    
    console.log('rowData preparada:', JSON.stringify(rowData));
    console.log('existingRow:', existingRow, 'lastRow:', lastRow);
    
    if (existingRow !== -1) {
      // Actualizar emisor existente
      const targetRow = existingRow + 2; // +2 porque empezamos en fila 2 y findIndex es 0-based
      // No actualizar la fecha de creación, mantener la existente
      const existingDate = emisoresSheet.getRange(targetRow, 10).getValue();
      rowData[9] = existingDate;
      
      console.log('Actualizando en fila:', targetRow);
      emisoresSheet.getRange(targetRow, 1, 1, 10).setValues([rowData]);
      console.log('Emisor actualizado:', emisorData.nombre);
    } else {
      // Añadir nuevo emisor
      console.log('Creando nuevo emisor en fila:', lastRow + 1);
      emisoresSheet.getRange(lastRow + 1, 1, 1, 10).setValues([rowData]);
      console.log('Nuevo emisor creado:', emisorData.nombre);
    }
    
    return {
      success: true,
      message: existingRow !== -1 ? 'Emisor actualizado' : 'Emisor creado'
    };
    
  } catch (error) {
    console.error('Error al guardar emisor:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Eliminar emisor
function deleteEmisor(emisorId) {
  try {
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let emisoresSheet = spreadsheet.getSheetByName('Emisores');
    
    if (!emisoresSheet) {
      return {
        success: false,
        error: 'La hoja de emisores no existe'
      };
    }
    
    const lastRow = emisoresSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: false,
        error: 'No hay emisores para eliminar'
      };
    }
    
    // Buscar el emisor por ID
    const ids = emisoresSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = ids.findIndex(id => id == emisorId); // Usar == para comparación flexible
    
    if (rowIndex === -1) {
      return {
        success: false,
        error: 'Emisor no encontrado'
      };
    }
    
    // Eliminar la fila (rowIndex + 2 porque empezamos en fila 2 y findIndex es 0-based)
    const targetRow = rowIndex + 2;
    emisoresSheet.deleteRow(targetRow);
    
    console.log('Emisor eliminado:', emisorId);
    
    return {
      success: true,
      message: 'Emisor eliminado correctamente'
    };
    
  } catch (error) {
    console.error('Error al eliminar emisor:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Reordenar filas según el orden actual del listado
function handleReorderRows(data) {
  try {
    const SPREADSHEET_ID = '1qCDvaMEERQ3lm1MLWQnl2TblwtxA-17hoFo8leG70kg';
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    
    const sortedData = data.data;
    const headers = data.headers;
    
    if (!sortedData || !Array.isArray(sortedData) || sortedData.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({
          result: 'error',
          error: 'No se proporcionaron datos válidos para reordenar'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Obtener el rango de datos actual (sin incluir cabeceras)
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // Limpiar datos existentes (excepto cabeceras)
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // Preparar los datos en el nuevo orden
    const newRows = sortedData.map(invoice => {
      return headers.map(header => invoice[header] || '');
    });
    
    // Insertar todos los datos de una vez para mejor rendimiento
    if (newRows.length > 0) {
      sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
    }
    
    console.log(`${newRows.length} filas reordenadas correctamente`);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'success',
        message: `${newRows.length} filas reordenadas correctamente`,
        rowsReordered: newRows.length
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error al reordenar filas:', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        result: 'error',
        error: error.toString(),
        message: 'Error al reordenar las filas'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}