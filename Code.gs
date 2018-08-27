function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Utilities').addItem('Pull Residuals', 'hyundaiResiduals').addToUi();
}
function hyundaiResiduals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('Table 1');
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var data = range.getDisplayValues();
  var storing = false;
  var initial = [];
  var terms = [36,39,42,48];
  var final = [];
  var residual, term, current;
  var found = false;
  var newSheet;
  
  for(var i=0;i<data.length;i++){
    if(data[i][1]=="Model#"){ storing = true; }
    if(storing){
      for(var j=0;j<data[i].length;j++){
        if(data[i][j]=="Residual"){ residual = j; j = data[i].length-1; i = data.length-1; }
      }
    }
  }
  storing=false;
  for(i=0;i<data.length;i++){
    found=false;
    if(data[i][1].split(" ")[0]=="MY"){
      term = parseInt(data[i][5].split(" ")[0]);
      for(j=0;j<terms.length;j++){
        if(terms[j]==term){ current = j+1; }
      }
    }
    if(data[i][1]=="Model#"){ storing = true; }
    else if(data[i][1]==""){ storing = false; }
    else if(storing){
      if(initial.length != 0) {
        for(j=0;j<initial.length&&!found;j++){
          if(initial[j][0]==data[i][1]){
            Logger.log("FOUND "+initial[j][0]);
            found=true;
            if(initial[j][current]!=0){
              var alert = ui.alert('ERROR', 'Initial array "'
                                   +initial[j]+'" already has a value stored for the '+term
                                   +' month term. Duplicate model code for a different year suspected. Would you like to retry with correction for duplicate model codes on separate years?', 
                                   ui.ButtonSet.YES_NO);
              if(alert == ui.Button.YES){ hyundaiResidualsYear(); }
              return;
            }
            initial[j][current]=data[i][residual];
          }
        }
      }
      if(!found){
        initial[initial.length]=[data[i][1]];
        for(j=0;j<terms.length;j++){
          initial[initial.length-1][j+1]=0;
        }
        initial[initial.length-1][current]=data[i][residual];
      }
    }
  }
  for(i=0;i<initial.length;i++){
    for(j=1;j<initial[i].length;j++){
      if(initial[i][j]==0){ initial[i][j] = ""; }
    }
  }
  final[0] = [""];
  for(i=0;i<terms.length;i++){final[0][i+1]=terms[i];}
  for(i=0;i<initial.length;i++){
    final[i+1]=initial[i];
  }
  if(ss.getSheetByName("Lease")!=null){ss.deleteSheet(ss.getSheetByName("Lease"));}
  newSheet = ss.insertSheet().setName("Lease");
  newSheet.deleteRows(final.length+1, newSheet.getMaxRows()-final.length);
  newSheet.deleteColumns(final[0].length+1, newSheet.getMaxColumns()-final[0].length);
  SpreadsheetApp.flush();
  newSheet.getRange(1, 1, final.length, final[0].length).setValues(final);
}

function hyundaiResidualsYear() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('Table 1');
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var data = range.getDisplayValues();
  var storing = false;
  var initial = [];
  var terms = [36,39,42,48];
  var final = [];
  var residual, term, current, year;
  var found = false;
  var newSheet;
  
  for(var i=0;i<data.length;i++){
    if(data[i][1]=="Model#"){ storing = true; }
    if(storing){
      for(var j=0;j<data[i].length;j++){
        if(data[i][j]=="Residual"){ residual = j; j = data[i].length-1; i = data.length-1; }
      }
    }
  }
  storing=false;
  for(i=0;i<data.length;i++){
    found=false;
    if(data[i][1].split(" ")[0]=="MY"){
      term = parseInt(data[i][5].split(" ")[0]);
      for(j=0;j<terms.length;j++){
        if(terms[j]==term){ current = j+2; }
      }
      year = parseInt(data[i][1].split(" ")[1]);
    }
    if(data[i][1]=="Model#"){ storing = true; }
    else if(data[i][1]==""){ storing = false; }
    else if(storing){
      if(initial.length != 0) {
        for(j=0;j<initial.length&&!found;j++){
          if(initial[j][1]==data[i][1] && initial[j][0]==year){
            Logger.log("FOUND "+initial[j][1]);
            found=true;
            if(initial[j][current]!=0){ ui.alert('ERROR', 'Initial array "'+initial[j]+'" already has a value stored for the '+term+' month term. Duplicate suspected. Halting operation.', ui.ButtonSet.OK); return; }
            initial[j][current]=data[i][residual];
          }
        }
      }
      if(!found){
        initial[initial.length]=[year,data[i][1]];
        for(j=0;j<terms.length;j++){
          initial[initial.length-1][j+2]=0;
        }
        initial[initial.length-1][current]=data[i][residual];
      }
    }
  }
  for(i=0;i<initial.length;i++){
    for(j=1;j<initial[i].length;j++){
      if(initial[i][j]==0){ initial[i][j] = ""; }
    }
  }
  final[0] = ["year","model"];
  for(i=0;i<terms.length;i++){final[0][i+2]=terms[i];}
  for(i=0;i<initial.length;i++){
    final[i+1]=initial[i];
  }
  if(ss.getSheetByName("Lease")!=null){ss.deleteSheet(ss.getSheetByName("Lease"));}
  newSheet = ss.insertSheet().setName("Lease");
  newSheet.deleteRows(final.length+1, newSheet.getMaxRows()-final.length);
  newSheet.deleteColumns(final[0].length+1, newSheet.getMaxColumns()-final[0].length);
  SpreadsheetApp.flush();
  newSheet.getRange(1, 1, final.length, final[0].length).setValues(final);
}
