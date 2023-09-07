class SpreadSheet{constructor({sheetId:t,sheetName:e,activeSheet:n=!0}){try{this.activeSpreadSheet=n?SpreadsheetApp.getActiveSpreadsheet():SpreadsheetApp.openById(t),this.activeSheet=this.activeSpreadSheet.getSheetByName(e),this.sheetId=n?this.activeSpreadSheet.getId():t,this.sheetName=e,this.sheetContent=[]}catch(s){this.activeSheet=null,this.sheetId=null,this.sheetName=null,this.sheetContent=s}}readAllContentDisplayFromSheet(){if(null==this.activeSheet)return{status:null,content:null};try{return this.sheetContent=this.activeSheet.getRange(1,1,this.activeSheet.getLastRow(),this.activeSheet.getLastColumn()).getDisplayValues(),{status:!0,content:this.sheetContent}}catch(t){return{status:!1,content:t}}}readAllContentValueFromSheet(){if(null==this.activeSheet)return{status:null,content:null};try{return this.sheetContent=this.activeSheet.getRange(1,1,this.activeSheet.getLastRow(),this.activeSheet.getLastColumn()).getValues(),{status:!0,content:this.sheetContent}}catch(t){return{status:!1,content:t}}}readValueFromACell(t){try{return{status:!0,content:this.activeSheet.getRange(t).getDisplayValue()}}catch(e){return{status:!1,content:e}}}writeValueOnACell(t,e){try{return this.activeSheet.getRange(t).setValue(e),{status:!0,content:!0}}catch(n){return{status:!1,content:n}}}writeValueOnARow(t,e){try{let n=this.activeSheet.getLastColumn();for(;e.length<n;)e.push("");return this.activeSheet.getRange(t,1,1,e.length).setValues([e]),{status:!0,content:!0}}catch(s){return{status:!1,content:s}}}writeValueOnAColumn(t,e){try{return e=e.flatMap(t=>[[t]]),this.activeSheet.getRange(1,t,e.length,1).setValues(e),{status:!0,content:!0}}catch(n){return{status:!1,content:n}}}readPartialContentFromSheet(t,e){try{return this.sheetContent=this.activeSheet.getRange(t,1,e,this.activeSheet.getLastColumn()).getDisplayValues(),{status:!0,content:this.sheetContent}}catch(n){return{status:!1,content:n}}}readRow(t){try{return this.sheetContent=this.activeSheet.getRange(t,1,1,this.activeSheet.getLastColumn()).getDisplayValues(),{status:!0,content:this.sheetContent}}catch(e){return{status:!1,content:e}}}readColumn(t){try{return this.sheetContent=this.activeSheet.getRange(1,t,this.activeSheet.getLastRow(),1).getDisplayValues(),{status:!0,content:this.sheetContent}}catch(e){return{status:!1,content:e}}}clearAll(t,e){try{return this.sheetContent=this.activeSheet.getRange(t,e,this.activeSheet.getLastRow(),this.activeSheet.getLastColumn()).clearContent(),{status:!0,content:this.sheetContent}}catch(n){return{status:!1,content:n}}}clearRow(t,e){try{return this.sheetContent=this.activeSheet.getRange(t,e,1,this.activeSheet.getLastColumn()).clearContent(),{status:!0,content:this.sheetContent}}catch(n){return{status:!1,content:n}}}clearColumn(t,e){try{return this.sheetContent=this.activeSheet.getRange(t,e,this.activeSheet.getLastRow(),1).clearContent(),{status:!0,content:this.sheetContent}}catch(n){return{status:!1,content:n}}}static acercaDe(){return{author:"Arturo Gomez => zesertebe@gmail.com",description:`Clase dise\xf1ada para trabajar hojas de calculo con el entorno AppScript proporcionando metodos quefacilitan la lectura y escritura de datos`,web:"https://adev.dev"}}writeValues(t,e){try{return this.activeSheet.getRange(t,1,e.length,e[0].length).setValues(e),{status:!0,content:!0}}catch(n){return{status:!1,content:n}}}}
