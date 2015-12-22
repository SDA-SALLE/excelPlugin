/*var form = form||{replace1:"B",with1:"A",column:16};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/Base_Fuentes_Industriales2.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/Base_Fuentes_Industriales_out.xlsx", {});*/
var file =  ExcelPOI.load("/tests/xx4.xlsx");
var data = file.copy(4,0,9,3);
//console.log(data);

file.refreshFormulas();

/*var form = form||{replace1:"CARBÓN VEGETAL",with1:"GAS",column:1};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/fijas/Base_Fuentes_Industriales2.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/fijas/Base_Fuentes_Industriales.xlsx", {});*/


/*var form = form||{replace1:3,with1:66.6,column:0};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/Libro1.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/Libro2.xlsx", {});*/

/*var form = form||{replace1:2,with1:5,column:5};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/base_fuentes_industriales_v1.xlsx");
//var baseFC = ExcelPOI.load("/tests/Base_Fuentes_Industriales2.xlsx");
baseFC.setCurrentSheet(3);
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
var result = baseFC.copy(5,4,6,8,{ format:"value",formulas:true})
//console.log(result);
baseFC.refreshFormulas();
baseFC.save("/tests/baseOut.xlsx", {});*/