function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Ventas Heidy Sofia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Función para manejar el inicio de sesión
function login(username, password) {
  if (!username || !password) {
    return null; // No permitir el inicio de sesión si los campos están vacíos
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mantenimiento');
  const data = sheet.getRange('A2:D').getValues(); // Columnas: A (Nombre), B (Usuario), C (Contraseña), D (Cargo)

  for (let i = 0; i < data.length; i++) {
    const [nombre, usuario, pass, cargo] = data[i];
    if (usuario.trim() === username.trim() && pass.trim() === password.trim()) {
      setEmpleadoLogueado(nombre); // Guardar el nombre del empleado logueado
      return { nombre, cargo };
    }
  }
  return null; // Retorna null si no encuentra coincidencia
}

function obtenerVendedores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mantenimiento');
  const vendedores = sheet.getRange('A2:A').getValues().flat().filter(String);
  return vendedores;
}

function obtenerClientes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CLIENTES');
  const clientes = sheet.getRange('B2:B').getValues().flat().filter(String);
  return clientes;
}

function buscarProducto(idProducto) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
  const data = sheet.getRange('B2:F').getValues();
  const productoEncontrado = data.find(row => row[0] == idProducto);

  if (productoEncontrado) {
    return {
      nombre: productoEncontrado[1],        // Columna C en hoja "Stock"
      talla: productoEncontrado[4],           // Columna F en hoja "Stock"
      precioPublico: productoEncontrado[2],   // Columna D en hoja "Stock"
      precioMayorista: productoEncontrado[3]  // Columna E en hoja "Stock"
    };
  } else {
    return null;
  }
}

// Función para recibir múltiples IDs y devolver la información correspondiente
function buscarProductos(ids) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
  const data = sheet.getRange('B2:F').getValues();
  var resultado = {};
  ids.forEach(function(idProducto) {
    var productoEncontrado = data.find(row => row[0] == idProducto);
    if (productoEncontrado) {
      resultado[idProducto] = {
        nombre: productoEncontrado[1],
        talla: productoEncontrado[4],
        precioPublico: productoEncontrado[2],
        precioMayorista: productoEncontrado[3]
      };
    } else {
      resultado[idProducto] = null;
    }
  });
  return resultado;
}

function buscarImagenProducto(nombreProducto) {
  // Se asume que la hoja PRODUCTOS contiene:
  // - Columna A: Nombre del producto
  // - Columna F: URL de la imagen o fórmula IMAGE("url")
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUCTOS');
  if (!sheet) return null;
  // Obtener tanto los valores como las fórmulas
  var data = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  // Se asume que la primera fila es encabezado; se recorre desde la fila 2 (índice 1)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == nombreProducto) { // Compara el nombre (columna A)
      // Primero, revisamos si existe una fórmula en la columna F (índice 5)
      var formula = formulas[i][5];
      if (formula && formula.indexOf("=IMAGE(") === 0) {
        // Extraer la URL de la fórmula IMAGE, que tiene la forma:
        // =IMAGE("https://drive.google.com/uc?export=view&id=FILEID")
        var urlMatch = formula.match(/=IMAGE\("([^"]+)"\)/);
        if (urlMatch && urlMatch[1]) {
          return urlMatch[1];
        } else {
          return null;
        }
      } else {
        // Si no es una fórmula, asumimos que la celda ya contiene la URL pura
        return data[i][5];
      }
    }
  }
  return null;
}

function registrarVentaWebApp(datosVenta, local, cliente, empleada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historialSheet = ss.getSheetByName('Historial_Ventas');
  var stockSheet = ss.getSheetByName('Stock');
  var stockVendidoSheet = ss.getSheetByName('STOCK_VENDIDO');

  if (!historialSheet || !stockSheet || !stockVendidoSheet) {
    return { success: false, message: "No se encontró alguna de las hojas necesarias" };
  }

  var fecha = new Date(); // Fecha y hora actual
  var lastRow = historialSheet.getLastRow();
  var newId = generarNuevoIdVenta(historialSheet);

  // Obtener datos actuales de Stock
  var stockData = stockSheet.getDataRange().getValues();
  var stockFormulas = stockSheet.getDataRange().getFormulas();

  var historialValues = [];
  var stockVendidoValues = [];
  var stockVendidoFormulas = [];
  var stockMap = {};

  stockData.forEach((row, i) => {
    stockMap[row[1]] = i + 1; // Mapea el ID de producto a la fila en 'Stock'
  });

  datosVenta.forEach(function(venta) {
    var idProducto = venta.idProducto;
    if (idProducto && stockMap[idProducto]) {
      historialValues.push([
        fecha,                // Columna A: Fecha
        empleada,             // Columna B: Empleada
        newId,                // Columna C: ID de Venta
        idProducto,           // Columna D: ID Producto
        venta.filaDatos[0],   // Columna E: Producto
        venta.filaDatos[1],   // Columna F: Talla
        venta.filaDatos[2],   // Columna G: Precio
        venta.pago,           // Columna H: Pago (valor global)
        venta.estado,         // Columna I: Estado (valor global)
        local,                // Columna J: Local de Venta
        cliente,              // Columna K: Cliente
        venta.comprobante     // Columna L: Comprobante N°
      ]);

      // Obtener la fila en 'Stock'
      var stockRowIndex = stockMap[idProducto] - 1;
      stockVendidoValues.push(stockData[stockRowIndex]);
      stockVendidoFormulas.push(stockFormulas[stockRowIndex]);

      // Marcar la fila para eliminación
      stockData[stockRowIndex] = null;
    }
  });

  // Actualizar Stock
  var newStockData = stockData.filter(row => row !== null);
  stockSheet.clear();
  if (newStockData.length > 0) {
    stockSheet.getRange(1, 1, newStockData.length, newStockData[0].length).setValues(newStockData);
  }

  // Insertar en Historial y Stock Vendido
  if (historialValues.length > 0) {
    historialSheet.getRange(lastRow + 1, 1, historialValues.length, historialValues[0].length).setValues(historialValues);
  }
  if (stockVendidoValues.length > 0) {
    var lastRowVendido = stockVendidoSheet.getLastRow();
    stockVendidoSheet.getRange(lastRowVendido + 1, 1, stockVendidoValues.length, stockVendidoValues[0].length).setValues(stockVendidoValues);
    stockVendidoFormulas.forEach(function(formulas, index) {
      formulas.forEach(function(formula, colIndex) {
        if (formula) {
          stockVendidoSheet.getRange(lastRowVendido + index + 1, colIndex + 1).setFormula(formula);
        }
      });
    });
  }

  return { success: true, message: "Venta registrada correctamente" };
}

function obtenerEmpleadoLogueado() {
  return PropertiesService.getScriptProperties().getProperty("EmpleadoLogueado");
}

function generarNuevoIdVenta(historialSheet) {
  var lastRow = historialSheet.getLastRow();
  var lastIdCell = historialSheet.getRange(lastRow, 3).getValue().toString(); // Columna C
  var lastIdNumber = parseInt(lastIdCell.split('-')[1]);
  var newIdNumber = lastIdNumber + 1;
  return 'VEN-' + ('00000000' + newIdNumber).slice(-8);
}

function setEmpleadoLogueado(nombre) {
  PropertiesService.getScriptProperties().setProperty("EmpleadoLogueado", nombre);
}

function cargarEmpleadas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mantenimiento');
  const empleadas = sheet.getRange('A2:A').getValues().flat().filter(String);
  return empleadas;
}

/* 
Función para cargar el contenido HTML de una página solicitada. 
Esta función se usará para la navegación dinámica desde el frontend.
*/
function getHtml(pageName) {
  return HtmlService.createHtmlOutputFromFile(pageName).getContent();
}
