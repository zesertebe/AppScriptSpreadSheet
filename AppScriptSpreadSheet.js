
/**
 * @class SpreadSheet
 * @author: Arturo G - zesertebe@gmail.com
 * @version: 1.0.8
 * @classdesc: Clase que proporciona metodos para leer y escribir datos en hojas de Google
 * se recomienda usar la extension de Chrome para una mejor visualizacion: https://chrome.google.com/webstore/detail/gase-google-appscript-edi/lefcemnilieamgifcegilmkaclmhakfc
 * */
class SpreadSheet {

  /**
   * ### *constructor*
   * Los parametros del constructor son los siguientes:
   * 
   * **✅ sheetId**: El identificador (id) de la hoja.
   * 
   * **✅ sheetName**: Nombre de la hoja
   * 
   * **✅ activeSheet**: Si es true(por defecto) significa que la hoja es la misma en donde se encuentra el proyecto. 
   * 
   * **Si la hoja no se pudo leer la propiedad activeSheet será igual a null**
   * 
   *     var spreadSheet = new SpreadSheet({
   *          sheetId: 'hf734hf943hd9j20t',
   *          sheetName: 'Hoja1',
   *          activeSheet: false
   *      })
   * 
   * 
   */
  constructor({ sheetId, sheetName, activeSheet = true }) {
    try {
      this.activeSpreadSheet = activeSheet ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(sheetId);
      this.activeSheet = this.activeSpreadSheet.getSheetByName(sheetName);
      this.sheetId = activeSheet ? this.activeSpreadSheet.getId() : sheetId;
      this.sheetName = sheetName;
      this.sheetContent = [];
    } catch (error) {
      this.activeSheet = null;
      this.sheetId = null;
      this.sheetName = null;
      this.sheetContent = error;
    }
  }

  /**
   * ### *readAllContentDisplayFromSheet*
   * Lee el contenido de toda la hoja.
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var sheetContent = spreadSheet.readAllContentDisplayFromSheet();
   * if(sheetContent.status != true){throw 'No es posible leer el contenido de la hoja'}
   * var arrayContent = sheetContent.content;
   * // obtenemos un arreglo que contiene los datos de la hoja
   * ```
   * 
   * Al ejectuar esta función el contendio de la hoja queda almacenado en el objeto **sheetContent** del prototipo instanciado de SpreadSheet
   * ```javascript
   * spreadSheet.readAllContentDisplayFromSheet();
   * var arrayContent = spreadSheet.sheetContent;
   * ```
   * 
   * @param  {void} - void
   * @return {Object} - objeto con la informacion leida de la hoja
   */
  readAllContentDisplayFromSheet() {
    if (this.activeSheet == null) {
      return { status: null, content: null };
    }
    else {
      try {
        this.sheetContent = this.activeSheet.getRange(1, 1, this.activeSheet.getLastRow(), this.activeSheet.getLastColumn()).getDisplayValues();
        return { status: true, content: this.sheetContent };
      } catch (error) {
        return { status: false, content: error }
      }
    }
  }

  /**
   * ### *readAllContentValueFromSheet*
   * Lee el contenido de toda la hoja.
   * 
   *     var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   *     var sheetContent = spreadSheet.readAllContentValueFromSheet();
   *     if(sheetContent.status != true){throw 'No es posible leer el contenido de la hoja'}
   *     var arrayContent = sheetContent.content;
   *     // obtenemos un arreglo que contiene los datos de la hoja
   * 
   * Al ejectuar esta función el contendio de la hoja queda almacenado en el objeto **sheetContent** del prototipo instanciado de SpreadSheet
   * 
   *     spreadSheet.readAllContentValueFromSheet();
   *     var arrayContent = spreadSheet.sheetContent;
   * @param  {void} - void
   * @return {Object} - objeto con la informacion leida de la hoja
   */
  readAllContentValueFromSheet() {
    if (this.activeSheet == null) {
      return { status: null, content: null };
    }
    else {
      try {
        this.sheetContent = this.activeSheet.getRange(1, 1, this.activeSheet.getLastRow(), this.activeSheet.getLastColumn()).getValues();
        return { status: true, content: this.sheetContent };
      } catch (error) {
        return { status: false, content: error }
      }
    }
  }

  /**
   * ### *readValueFromACell*
   * 
   * Lee el contenido de una sola celda
   * 
   *     var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   *     var cellContent = spreadSheet.readValueFromACell('A1');
   *     if(cellContent.status != true){throw 'No es posible leer el contenido de la hoja'}
   *     var content = cellContent.content;
   *     // obtenemos el valor de esa celda
   * 
   * @param {string} cell Celda a leer en formato A1
   * @return {Object} {} - objeto con la informacion leida de la hoja
   */
  readValueFromACell(cell) {
    try {
      return { status: true, content: this.activeSheet.getRange(cell).getDisplayValue() };
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Escribe datos en una sola celda
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.writeValueOnACell('A1', 'Texto');
   * if(result.status != true){throw 'No es posible escribir el contenido de la hoja'}
   * ```
   * 
   * @param {string} cell Celda para escribir en formato A1
   * @param {string} data Datos para escribir en la velda
   * @return {Object} {} - objeto que informa del estado
   */
  writeValueOnACell(cell, data) {
    try {
      this.activeSheet.getRange(cell).setValue(data);
      return { status: true, content: true }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Escribe datos en una sola fila
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.writeValueOnARow(1, ['texto', 3, 4]);
   * if(result.status != true){throw 'No es posible escribir el contenido de la hoja'}
   * ```
   * 
   * @param {number} row Fila en la cual queremos escribir datos
   * @param {Object} dataArray Arreglo de datos para escribir en la fila
   * @return {Object} {} - objeto que informa del estado
   */
  writeValueOnARow(row, dataArray) {
    try {
      this.activeSheet.getRange(row, 1, 1, this.activeSheet.getLastColumn()).setValues([dataArray]);
      return { status: true, content: true };
    } catch (error) {
      return { status: false, content: error };
    }
  }

  /**
   * ### Escribe datos en una sola columna
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.writeValueOnAColumn(1, [['Texto', 1, 2]]);
   * if(result.status != true){throw 'No es posible escribir el contenido de la hoja'}
   * ```
   * 
   * @param {number} column Numero de la columna en donde escribir los datos
   * @param {Object} dataArray EL arreglo de datos para escribir en la columna
   * @return {Object} {} - objeto que informa del estado
   */
  writeValueOnAColumn(column, dataArray) {
    try {
      dataArray = dataArray.flatMap(el => [[el]])
      this.activeSheet.getRange(1, column, dataArray.length, 1).setValues(dataArray);
      return { status: true, content: true };
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Lee solamente una parte de la hoja
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.readPartialContentFromSheet(1, 4);
   * // lee solamente las primeras 4 filas
   * if(result.status != true){throw 'No es posible leer el contenido de la hoja'}
   * ```
   * 
   * @param {number} start Celda para escribir en formato A1
   * @param {number} end Datos para escribir en la velda
   * @return {Object} {} - objeto que informa del estado
   */
  readPartialContentFromSheet(start, end) {
    try {
      this.sheetContent = this.activeSheet.getRange(start, 1, end, this.activeSheet.getLastColumn()).getDisplayValues();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Lee solamente una fila
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.readRow(3);
   * if(result.status != true){throw 'No es posible escribir el contenido de la hoja'}
   * ```
   * 
   * @param {number} row La fila que queremos leer
   * @return {Object} {} - objeto que informa del estado
   */
  readRow(row){
    try {
      this.sheetContent = this.activeSheet.getRange(row, 1, 1, this.activeSheet.getLastColumn()).getDisplayValues();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Lee solamente una columna
   * 
   * ```javascript
   * var spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * var result = spreadSheet.readColumn(3);
   * if(result.status != true){thcolumn 'No es posible escribir el contenido de la hoja'}
   * ```
   * 
   * @param {number} column La columna que queremos leer
   * @return {Object} {} - objeto que informa del estado
   */
  readColumn(column){
    try {
      this.sheetContent = this.activeSheet.getRange(1, column, this.activeSheet.getLastRow(), 1).getDisplayValues();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Limpia todos los datos en el rango especificado
   * 
   * ```javascript
   * let spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * let result = spreadSheet.clearAll(1, 2);
   * // elimina todos los datos desde la fila 1 y columna 2
   * if(result.status != true){throw 'No es posible eliminar los datos'}
   * ```
   * 
   * @param {number} row La fila desde la cual queremos realizar la limpieza
   * @param {number} column La columna desde la cual queremos realizar la limpieza
   * @return {Object} {} - objeto que informa del estado
   */
  clearAll(row, column) {
    try {
      this.sheetContent = this.activeSheet.getRange(row, column, this.activeSheet.getLastRow(), this.activeSheet.getLastColumn()).clearContent();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Limpia todos los datos en una sola fila
   * 
   * ```javascript
   * let spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * let result = spreadSheet.clearRow(1, 2);
   * // elimina todos los datos de la fila 1 empezando en la columna 2
   * if(result.status != true){throw 'No es posible eliminar los datos'}
   * ```
   * 
   * @param {number} row La fila que queremos limpiar
   * @param {number} column La columna desde la cual queremos realizar la limpieza
   * @return {Object} {} - objeto que informa del estado
   */
  clearRow(row, column) {
    try {
      this.sheetContent = this.activeSheet.getRange(row, column, 1, this.activeSheet.getLastColumn()).clearContent();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ### Limpia todos los datos en una sola columna
   * 
   * ```javascript
   * let spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * let result = spreadSheet.clearColumn(1, 2);
   * // elimina todos los datos de la columna 2 empezando desde la fila 1
   * if(result.status != true){throw 'No es posible eliminar los datos'}
   * ```
   * 
   * @param {number} row La fila desde la cual queremos realizar la limpieza
   * @param {number} column La columna que queremos limpiar
   * @return {Object} {} - objeto que informa del estado
   */
  clearColumn(row, column) {
    try {
      this.sheetContent = this.activeSheet.getRange(row, column, this.activeSheet.getLastRow(), 1).clearContent();
      return { status: true, content: this.sheetContent }
    } catch (error) {
      return { status: false, content: error }
    }
  }

  /**
   * ## Desarrollado por:
   * 
   * 
   * ### Arturo Gomez => zesertebe@gmail.com
   * 
   * ### Visite: [ocancelada.dev](https://ocancelada.dev)
   * 
   * ### se recomienda usar la extension de Chrome para una mejor visualizacion: [GASE](https://chrome.google.com/webstore/detail/gase-google-appscript-edi/lefcemnilieamgifcegilmkaclmhakfc)
   * 
   * ```mermaid
   * Clase diseñada para trabajar hojas de calculo con el entorno AppScript proporcionando metodos que 
   * facilitan la lectura y escritura de datos.
   * ```
   * 
   * > *2022*
   * 
   */
  static acercaDe() {
    return {
      author: 'Arturo Gomez => zesertebe@gmail.com',
      description: `Clase diseñada para trabajar hojas de calculo con el entorno AppScript proporcionando metodos quefacilitan la lectura y escritura de datos`,
      web: 'https://ocancelada.dev',
    }
  }

  /**
   * ### Escribir datos en varias filas
   * 
   * ```javascript
   * let spreadSheet = new SpreadSheet({sheetName: 'Hoja1', activeSheet: true});
   * let data = [['Pepe', 123456789, 'Activo'], ['Juania', 666666789, 'Activo'],['Karla', 666666666, 'Inactivo']]
   * let result = spreadSheet.writeValues(1, data);
   * // Escribe los datos empezando desde la fila 1
   * if(result.status != true){throw 'No es posible escribir los datos'}
   * ```
   * 
   * @param {number} row La fila desde la cual queremos escribir los datos
   * @param {object} data EL arreglo de datos que queremos escribir
   * @return {Object} {} - objeto que informa del estado
   */
  writeValues(row, data) {
    try {
      this.activeSheet.getRange(row, 1, data.length, data[0].length).setValues(data);
      return { status: true, content: true }
    } catch (error) {
      return { status: false, content: error }
    }
  }
}

