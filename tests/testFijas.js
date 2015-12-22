
/*var form = form||{replace1:3,with1:66.6,column:0};
//console.log(form);
var baseFC = ExcelPOI.load("/tests/Libro1.xlsx");
//console.log(baseFC);
baseFC.findAndReplaceCol(form.column, form.replace1, form.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/Libro2.xlsx", {});
*/

//var form = form||{replace1:2,with1:5,column:5};
/*var inputs = {replace1:"B",with1:"C",column:15};
var baseFC = ExcelPOI.load("/tests/fijas/Base_FIA.xlsx");
baseFC.findAndReplaceCol(inputs.column, inputs.replace1, inputs.with1);
baseFC.refreshFormulas();
baseFC.save("/tests/fijas/Base_FIA_out.xlsx", {});*/

var rows = [];
for(var i=2;i<468;i++)
  rows[i-2]=i;
var columns = [32,39,43,48,50];

var baseFC = ExcelPOI.load("/tests/fijas/Base_FCA.xlsx");//ExcelPOI.load(Global.currentDir+"input/Base_ESA.xlsx");
baseFC.setCurrentSheet(0);
baseFC.refreshFormulas();
var emisionesSinCambios = JSON.parse(baseFC.copy(1,58,845,64,{format:"value"}));

console.log(emisionesSinCambios[0][0]);

