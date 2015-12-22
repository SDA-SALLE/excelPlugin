/**
 * @object ExcelPOI
 * Library that provides methods for excel manipulation through the POI library.
 * @constructor
 * Load a new excel file
 * @param	filename:string	The path of the excel file
 * @return	+ExcelPOI
 */
var ExcelPOI = function (newExcelWB) {
	this.ExcelWB=newExcelWB;
};

/**
 * @function load(filename)
 * This function loads and returns an Excel workbook.
 * @param	filename:string	The path of excel file
 * @return	+ExcelPOI	The loaded workbook
 * 
 */
ExcelPOI.load = function(filename) {
	return new ExcelPOI(ExcelPOIAPI.load(Global.basedir, Global.basedirkey, File.checkGlobal(filename)));
};

/**
 * @function create(data)
 * This function creates a new Excel workbook. The workbook can be initialized by using the
 * a JSON object. Not clear how to do it right now.
 * @param 	data:+Object	A JSON object containing the workbook information.
 * @return	+ExcelPOI	The new workbook.
 */
ExcelPOI.create=function(data) {
	return new ExcelPOI(ExcelPOIAPI.create(Global.basedir, Global.basedirkey, data));
};

/**
 * @object ExcelPOI.prototype
 * Methods of the ExcelPOI object
 */
ExcelPOI.prototype = {
		/**
		 * @function	save(path, options)
		 * Saves the given excel workbook in the format specified by the extension of the
		 * path. The allowed extensions are: xlsx and xls
		 * @param 		path:string		physical path in which to save the image
		 * @param		options:+Object	Object containing the options
		 * @return 		string		the given read URL.
		 */
		save: function(path, options) {
			return File.getReadURL(this.ExcelWB.save(File.checkGlobal(path), options));
		},
		
		/**
		 * @function refreshFormulas()
		 * Refresh all the cell formulas in the workbook.
		 */
		refreshFormulas: function(){
			this.ExcelWB.refreshFormulas();
		},
		
		/**
		 * @function setCurrentSheet(index)
		 * To set the current index. 
		 * @param 	index:number	Index of the sheet to set
		 */
		setCurrentSheet: function(index) {
			this.ExcelWB.setCurrentSheet(index);
		},
		
		/**
		 * @function getCurrentSheet()
		 * To get the current sheet.
		 * @return +Object The current XSSFSheet
		 */
		getCurrentSheet: function() {
			return this.ExcelWB.getCurrentSheet();
		},
		
		/**
		 * @function findAndReplaceCol(column, from, to)
		 * Find all the occurrences of 'from' in the given column and replace them by 'to'.
		 * @param column:number The column to use
		 * @param from:string The value to find. String or number
		 * @param value:string The value to replace. String or number   
		 * @return bool true if any replacing was made.
		 */
		findAndReplaceCol: function(column, from, to){
			return this.ExcelWB.findAndReplaceCol(column, from, to);
		},
		
		/**
		 * @function findAndReplaceRow(row, from, to)
		 * Find all the occurrences of 'from' in the given row and replace them by 'to'.
		 * @param row:number The row to use
		 * @param from:string The value to find. String or number
		 * @param value:string The value to replace. String or number  
		 * @return bool true if any replacing was made.
		 */
		findAndReplaceRow:function(row, from, to){
			return this.ExcelWB.findAndReplaceRow(row, from, to);
		},
		
		/**
		 * @function findAndReplaceAll(from, to)
		 * Find all the occurrences of 'from' in the current sheet and replace them by 'to' 
		 * @param from:string The value to find. String or number
		 * @param value:string The value to replace. String or number 
		 * @return bool true if any replacing was made.
		 */
		findAndReplaceAll: function(from, to){
			return this.ExcelWB.findAndReplaceAll(from, to);
		},
		
		/**
		 * @function findInRow(row, value)
		 * It finds the first occurrence of 'value' in the given row and returns the column index.
		 * @param row:number The row to use
		 * @param value:string The value to find 
		 * @return number The column index
		 */
		findInRow: function(row,value){
			return this.ExcelWB.findInRow(row, value);
		},
		
		/**
		 * @function findInColumn(column, value)
		 * It finds the first occurrence of 'value' in the given column and returns the row index 
		 * @param column:number The column to use
		 * @param value:number The value to find
		 * @return number The row index
		 */
		findInColumn:function(column, value){
			return this.ExcelWB.findInColumn(column, value);
		},
		
		/**
		 * @function paste(row, column,   matrixObject)
		 * It pastes the given cells in the current sheet starting at (row,col) cell.
		 * @param row:number
		 * @param column:number
		 * @param matrixObject:+Object
		 * @return bool false if it fails at copying the values.
		 */
		paste: function(column, row,  matrixObject){
			return this.ExcelWB.paste(row, column, matrixObject);
		},
		
		/**
		 * @function copy(row0, column0, row1, column1, options)
		 * It copies the given cells of the current sheet.
		 * @param row0:number
		 * @param column0:number
		 * @param row1:number
		 * @param column1:number
		 * @option format:string ["cell"|"json"|"value"] 
		 * @return +Object A JSONArray of XSSFCell
		 */
		copy: function(row0, column0, row1, column1, options){
			return this.ExcelWB.copy(row0, column0, row1, column1, options);
		},
		
		/**
		 * @function copyRanges(rows, columns, options)
		 * It copies the given cells of the current sheet.
		 * @param rows:Object
		 * @param columns:Object
		 * @option format:string ["cell"|"json"|"value"] 
		 * @return +Object A JSONArray of XSSFCell
		 */
		copyRanges: function(rows,columns, options){
			return this.ExcelWB.copy(rows, columns, options);
		},	
		
		/**
		 * @function put(row, column, cellObject)
		 * Sets the given cell to the given value.
		 * @param rows:Object
		 * @param columns:Object
		 * @option cellObject:Object 
		 * @example data.put(0,0,{type:"string",value:"hello!"});
		 */
		put: function(row,column, cellObject){
			return this.ExcelWB.put(row, column, cellObject);
		},
		
		/**
		 * @function nRows(column,startIndex)
		 * Returns the index maximum index of the last non null element 
		 * @param column:Number The column to check
		 * @param startIndex:Number The index to start the look up[Default 0]
		 */
		nRows: function(column,startIndex){
			if(startIndex)
				return this.ExcelWB.nRows(column,startIndex);
			else
				return this.ExcelWB.nRows(column,0);
		}
};