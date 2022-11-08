# AppScriptSpreadSheet
Proporciona metodos para leer y escribir datos en hojas de Google

> puede incluir esta librería en cualquier proyecto de tipo AppScript

## Ejemplos:

```javascript
// Instanciar la clase SpreadSheet:
let sheet = new SpreadSheet({
	sheetId: 'SHEET_ID', // identificador de la hoja
	sheetName: 'NAME_SHEET', // nombre de la hoja
	activeSheet: false // si es true significa que la hoja es la misma en donde se encuentra el script.
	// en ese caso el parametro sheetId no es necesario
})

// leer TODOS los datos de la hoja;
sheet.readAllContentDisplayFromSheet();

// limpiar la primera fila de la hoja desde la 2 columna (Columna B):
sheet.clearRow(1,2)

// escribir en una sola celda (A3)
sheet.writeValueOnACell('A3');

// escribir datos en una sola columna (Columna C)
sheet.writeValueOnAColumn(3, [ ['dato1', 'datao2', 'dato3'] ])

// leer solamente los datos en la columna H
sheet.readColumn(8);
```

## Métodos

| # | Nombre | Parámetros | Descripción
-- | -- | -- | -- |
| 1 | <strong style="color: #CBC700;">readAllContentDisplayFromSheet</strong> | void | Lee el contenido disponible de toda la hoja. Retorna un objeto con la información. Al mismo guarda esa informacion en la propiedad sheetContent de la instanciación
| 2 | <strong style="color: #CBC700;">readAllContentValueFromSheet</strong> | void | Lee los valores de toda la hoja. Retorna un objeto con la información. Al mismo guarda esa informacion en la propiedad sheetContent de la instanciación
| 3 | <strong style="color: #CBC700;">writeValueOnACell</strong> | cell:string, data:any | Escribe información en una sola celda.
| 4 | <strong style="color: #CBC700;">writeValueOnARow</strong> | row: number, dataArray:Array | Escribe un arreglo de datos en una sola fila
| 5 | <strong style="color: #CBC700;">writeValueOnAColumn</strong> | column: number, dataArray; Array | Escribe un arreglo de datos en una sola columna
| 6 | <strong style="color: #CBC700;">readPartialContentFromSheet</strong> | start:number, end:number | Lee solamente una parte de la hoja Retorna un objeto con la información. Al mismo guarda esa informacion en la propiedad sheetContent de la instanciación
| 7 | <strong style="color: #CBC700;">readRow</strong> | row:number | Lee solamente una fila. Retorna un objeto con la información. Al mismo guarda esa informacion en la propiedad sheetContent de la instanciación
| 8 | <strong style="color: #CBC700;">readColumn</strong> | column:string |  Lee el contenido de una sola columna. Retorna un objeto con la información. Al mismo guarda esa informacion en la propiedad sheetContent de la instanciación
| 9 | <strong style="color: #CBC700;">clearAll</strong> | row:number, column:number | Elimina el contenido de toda la hoja desde la fila y columna especificadas
| 10 | <strong style="color: #CBC700;">clearRow</strong> | row:number, column:number | Limpia una sola fila desde la columna especificada
| 11 | <strong style="color: #CBC700;">clearColumn</strong> | row:number, column:number |  Limpia una sola columna desde la fila especificada
| 12 | <strong style="color: #CBC700;">writeValues</strong> | row:number, data:Array |  Escribe datos en la hoja desde la fila especificada
