function Vaia() {
  // Set a comment on the edited cell to indicate when it was changed.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet=ss.getSheetByName("Movimenti");
  var NumR=sheet.getActiveCell().getRow();
  var NumC=sheet.getActiveCell().getColumn();
   
  if((NumR==1)&&(NumC==12)) {
    var range=ss.getRange("A200");
    sheet.setActiveRange(range);
    
    
  } 
}


function Resetta() {
   var sh = SpreadsheetApp.getActiveSheet(); 
  
  sh.getRange("L8:T8").clear();
}

function Modifica(NumRDest) {
  var sh = SpreadsheetApp.getActiveSheet(); 
  
//   var NumRDest=sh.getRange("W5").getValue();
  Logger.log("1 "+sh.getRange("L6:T6").getValues()+" NumR "+NumRDest);
      
   var   sheetDest=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movimenti");
  sheetDest.getRange(NumRDest,1,1,9).setValues(sh.getRange("L6:T6").getValues());
  Resetta();
  
      }  

function Inserimento() {
   var ID_C=SpreadsheetApp.getActiveSheet().getRange("L1").getValue();
       var impor=SpreadsheetApp.getActiveSheet().getRange("D2").getValue();
        var sh = SpreadsheetApp.getActiveSheet(); 
       var ID_CAT= sh.getRange("E2").getValue();
       var nump=ID_C*impor*ID_CAT; 
       Logger.log("NUmP="+nump);
      if(nump!=0) { 
         var NumOper= sh.getRange("L2").getValue();
         var NumRDest=sh.getRange("L1").getValue();
        var   sheetDest=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movimenti");
   
        var range=sh.getRange(33,1,NumOper,9).getValues();
        var currentRow = sheetDest.getLastRow();
      var sourceRange = sheetDest.getRange(currentRow, 10);
         var sourceFormulas = sourceRange.getFormulasR1C1();
        for(i=0;i<NumOper;i++)  {
        
          sheetDest.appendRow(range[i]);
          currentRow++;
         var targetRange =sheetDest.getRange(currentRow, 10);
         targetRange.setFormulasR1C1(sourceFormulas);
        }  
      //  sheetDest.getRange(NumRDest,1,NumOper,9).setValues(sh.getRange(33,1,NumOper,9).getValues());
       
   
          sh.getRange("A2:J2").clear();     
      }  
}  


function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
  if (!e || e.value === undefined)     return;
  const edited = e.range;
  const ss = edited.getSheet();
  var sh = SpreadsheetApp.getActiveSheet();
   var NumR=sh.getActiveCell().getRow();
     //  sheet.getActiveCell().getRow();
     var NumC=sh.getActiveCell().getColumn();
  Logger.log("NumC="+NumC+" NumR="+NumR);
    var s = SpreadsheetApp.getActiveSpreadsheet();
      var   sheetDest=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movimenti");
     Logger.log("sheet="+ss.getName());
  
    if (ss.getName() == "Ricerca"){
      
      if(((NumC==23) && (NumR==5)) || ((NumC==23) && (NumR==5))){
         Logger.log("NumC="+NumC," NumR="+NumR);
         sh.getRange("L8:T8").setValues(sh.getRange("B6:J6").getValues());
      }
      if((NumC==23) && (NumR==8)) {
         var currentRow = sheetDest.getLastRow();
      var sourceRange = sheetDest.getRange(currentRow, 10);
         var sourceFormulas = sourceRange.getFormulasR1C1();
        
        Modifica(sh.getRange("T3").getValue());
       currentRow++;
         var targetRange =sheetDest.getRange(currentRow, 10);
         targetRange.setFormulasR1C1(sourceFormulas);
      }
      if((NumC==21) && (NumR==8)) {
        Modifica(sh.getRange("W5").getValue());
      } 
   //   if((NumC==23) && (NumR==8)) Resetta();
       if((NumC==23) && (NumR>9)) {
        var ID = sh.getRange(NumR,NumC-1).getValue();
        sh.getRange("W5").setValue(ID);
         sh.getRange("L8:T8").setValues(sh.getRange("B6:J6").getValues());
         sh.getRange("W10:W199").setValues(sh.getRange("X10:X199").getValues());
      }  
      
      if((NumC==22) && (NumR==8)) {
         var NumRDest=sh.getRange("W5").getValue();
         Logger.log("NumRDest="+NumRDest);      
        sheetDest.deleteRow(NumRDest);
        sh.getRange("L8:T8").clear();
     
      } 
   
    }  
  
  
  if (ss.getName() == "Inserimento") {
      Logger.log("NumC="+NumC+" NumR="+NumR);
   
    if((NumC==11) && (NumR==2))    Inserimento();
    
  }  
  
  
 
  
  }

