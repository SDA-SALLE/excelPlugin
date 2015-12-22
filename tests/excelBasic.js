/*var form = form||{replace1:"CARBÓN VEGETAL",with1:"GAS",column:15};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/fijas/Base_Fuentes_comerciales_v2.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/fijas/Base_Fuentes_comerciales_out.xlsx", {});
*/

/*var form = form||{replace1:3,with1:66.6,column:0};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/Libro1.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/Libro2.xlsx", {});+*/

//var form = form||{replace1:2,with1:5,column:5};
var rows = [];
for(var i=2;i<498;i++)
  rows[i-2]=i;
var columns = [32,39,43,48,50];

var baseFC = ExcelPOI.load(Global.currentDir+"input/Base_ESA.xlsx");
baseFC.setCurrentSheet(0);
var emisionesSinCambios = JSON.parse(baseFC.copy(rows, columns, {format:"value"}));



/*baseFC.setCurrentSheet(3);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
var result = baseFC.copy(5,4,6,8,{ format:"value",formulas:true})
baseFC.save("/tests/baseOut.xlsx", {});*/

/*
var form = form||{replace1:2,with1:5,column:5};
//console.log(form);
//var baseFC = ExcelPOI.load("/tests/base_fuentes_industriales_v1.xlsx");
var baseFC = ExcelPOI.load("/tests/Base_Fuentes_Industriales2.xlsx");
baseFC.setCurrentSheet(3);
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
var result = baseFC.copy(5,4,6,8,{ format:"value",formulas:true})
//console.log(result);
baseFC.refreshFormulas();
baseFC.save("/tests/baseOut.xlsx", {});*/