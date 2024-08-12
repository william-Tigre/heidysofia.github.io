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
    const data = sheet.getRange('A2:D').getValues(); // Columna A (Nombre), B (Usuario), C (Contraseña), D (Cargo)

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
    const vendedores = sheet.getRange('A2:A').getValues().flat().filter(String); // Obtener valores no vacíos
    return vendedores;
}

function obtenerClientes() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CLIENTES');
    const clientes = sheet.getRange('B2:B').getValues().flat().filter(String); // Obtener valores no vacíos
    return clientes;
}

function buscarProducto(idProducto) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock');
    const data = sheet.getRange('B2:F').getValues(); // Asegúrate de que las columnas son B:F
    const productoEncontrado = data.find(row => row[0] == idProducto);

    if (productoEncontrado) {
        return {
            nombre: productoEncontrado[1],       // Columna C en hoja "Stock"
            talla: productoEncontrado[4],        // Columna F en hoja "Stock"
            precioPublico: productoEncontrado[2],// Columna D en hoja "Stock"
            precioMayorista: productoEncontrado[3] // Columna E en hoja "Stock"
        };
    } else {
        return null; // Producto no encontrado
    }
}

function registrarVentaWebApp(datosVenta, local, cliente) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var historialSheet = ss.getSheetByName('Historial_Ventas');
    var stockSheet = ss.getSheetByName('Stock');
    var stockVendidoSheet = ss.getSheetByName('STOCK_VENDIDO');
    var empleado = obtenerEmpleadoLogueado(); // Función que devuelve el nombre del empleado que inició sesión

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
        stockMap[row[1]] = i + 1; // Crear un mapa de ID de productos a su fila en 'Stock'
    });

    datosVenta.forEach(function(venta) {
        var idProducto = venta.idProducto;
        if (idProducto && stockMap[idProducto]) {
            historialValues.push([
                fecha,                          // Columna A: Fecha
                empleado,                       // Columna B: Empleada
                newId,                          // Columna C: ID de Venta
                idProducto,                     // Columna D: ID Producto
                venta.filaDatos[0],             // Columna E: Producto
                venta.filaDatos[1],             // Columna F: Talla
                venta.filaDatos[2],             // Columna G: Precio
                venta.filaDatos[3],             // Columna H: Pago
                venta.filaDatos[4],             // Columna I: Estado
                local,                          // Columna J: Local de Venta
                cliente                         // Columna K: Cliente
            ]);

            // Obtener la fila correspondiente en 'Stock'
            var stockRowIndex = stockMap[idProducto] - 1;
            stockVendidoValues.push(stockData[stockRowIndex]);
            stockVendidoFormulas.push(stockFormulas[stockRowIndex]);

            // Marcar la fila para eliminación
            stockData[stockRowIndex] = null; // Marcar la fila como eliminada
        }
    });

    // Eliminar filas en 'Stock'
    var newStockData = stockData.filter(row => row !== null);
    stockSheet.clear(); // Limpiar la hoja antes de volver a escribir los datos
    if (newStockData.length > 0) {
        stockSheet.getRange(1, 1, newStockData.length, newStockData[0].length).setValues(newStockData);
    }

    // Insertar en 'Historial_Ventas' y 'STOCK_VENDIDO'
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
    var lastIdCell = historialSheet.getRange(lastRow, 3).getValue().toString(); // Columna C en la última fila
    var lastIdNumber = parseInt(lastIdCell.split('-')[1]); // Extraer el número del ID
    var newIdNumber = lastIdNumber + 1;
    return 'VEN-' + ('00000000' + newIdNumber).slice(-8); // Formatear el nuevo ID
}

function setEmpleadoLogueado(nombre) {
    PropertiesService.getScriptProperties().setProperty("EmpleadoLogueado", nombre);
}
