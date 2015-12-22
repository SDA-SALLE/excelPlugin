package org.cheminfo.scripting.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.cheminfo.function.scripting.SecureFileManager;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
/**
 * 
 * @author acastillo
 *
 */

public class ExcelWorkbook extends XSSFWorkbook{
	private String basedir;
	private String key;
	private XSSFSheet currentSheet;
	private ExcelPOI caller;
	
	public ExcelWorkbook(String basedir, String key, String filename, ExcelPOI excelFunction) throws FileNotFoundException, IOException{
		super(new FileInputStream(filename));
		this.caller = excelFunction;
		this.basedir = basedir;
		this.key = key;
		this.setActiveSheet(0);
		this.currentSheet = this.getSheetAt(0);
	}
	
	public ExcelWorkbook(String basedir, String key, Object infoObject, ExcelPOI excelFunction) throws FileNotFoundException, IOException{
		super();
		this.caller = excelFunction;
		this.basedir = basedir;
		this.key = key;
		JSONObject info = caller.checkParameter(infoObject);
		String[] names = JSONObject.getNames(info);
		try {
			for(String name:names){
				JSONObject sheet = info.getJSONObject(name);
				if(sheet.has("type")&&sheet.getString("type").compareTo("sheet")==0){
					currentSheet = this.createSheet();
				}
			}
		} catch (JSONException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	/**
	 * Save the content of this ExcelWorkbook in the given file.
	 * @param path
	 * @return
	 */
	public String save(String path) {
		return save(path, null);
	}
	
	/**
	 * Saves the given Excel file in the format specified by the extension of the
	 * path.
	 * 
	 * @param image
	 * @param fullName
	 * @param options
	 *            {quality:(0-100)} only works for xlsx
	 * @return String: The name of the saved file. 
	 */
	public String save(String name, Object options) {
		try {
			String fullName = SecureFileManager.getValidatedFilename(basedir, key, name);
			if (fullName == null) {
				caller.appendError("Excel::save", "The file path is null");
				return "";
			}
			SecureFileManager.mkdir(basedir, key, name.replaceAll("[^/]*$", ""));
			int dotLoc = fullName.lastIndexOf('.');
			String format = fullName.substring(dotLoc + 1);
			format = format.toLowerCase(Locale.US);

			FileOutputStream dataFileOut;
			dataFileOut = new FileOutputStream(fullName);
			this.write(dataFileOut);
			dataFileOut.flush();
			dataFileOut.close();
			
		} catch (Exception ex) {
			caller.appendError("Excel::save", "Error : " + ex.toString());
			return "";
		}
		return name;
	}
	
	/**
	 * Recalculates all the formulas on the workbook
	 */
	public void refreshFormulas(){
		XSSFFormulaEvaluator.evaluateAllFormulaCells(this);
	}
	
	/**
	 * Set the current sheet
	 * @param i
	 */
	public void setCurrentSheet(int i){
		this.setActiveSheet(i);
		currentSheet = this.getSheetAt(i);
	}
	
	/**
	 * Return the current sheet
	 * @return XSSFSheet
	 */
	public XSSFSheet getCurrentSheet(){
		return currentSheet;
	}
	
	/**
	 * This function searches in the specified column for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param column
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceCol(int column, String from, String to){
		from = from.replaceAll("[^\\x00-\\x7F]", "");
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		boolean toReturn = false;
		//System.out.println(this.currentSheet);
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			if(row.getCell(column)!=null){
				try{
					if(row.getCell(column).getStringCellValue().replaceAll("[^\\x00-\\x7F]", "").compareTo(from)==0){
						row.getCell(column).setCellValue(to);
						toReturn=true;
					}
				}catch(IllegalStateException e){
					//Do nothing. Just ignore this cell
				}
					
			}
		}
		return toReturn;
		//return false;
	}
	
	/**
	 * This function searches in the specified column for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param column
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceCol(int column, double from, double to){
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		boolean toReturn = false;
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			if(row.getCell(column)!=null){
				try{
					if(row.getCell(column).getNumericCellValue()==from){
						row.getCell(column).setCellValue(to);
						toReturn=true;
					}
				}catch(IllegalStateException e){
					//Do nothing. Just ignore this cell
				}
			}
		}
		return toReturn;
	}
	
	/**
	 * This function searches in the specified row for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param row
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceRow(int row, String from, String to){
		from = from.replaceAll("[^\\x00-\\x7F]", "");
		XSSFRow rowCells = this.currentSheet.getRow(row);
		Iterator<Cell> cellIterator = rowCells.cellIterator();
		Cell cell = null;
		boolean toReturn = false;
		while(cellIterator.hasNext()){
			cell = cellIterator.next();
			try{
				if(cell.getStringCellValue().replaceAll("[^\\x00-\\x7F]", "").compareTo(from)==0){
					cell.setCellValue(to);
					toReturn=true;
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return toReturn;
	}
	
	/**
	 * This function searches in the specified row for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param row
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceRow(int row, double from, double to){
		XSSFRow rowCells = this.currentSheet.getRow(row);
		Iterator<Cell> cellIterator = rowCells.cellIterator();
		Cell cell = null;
		boolean toReturn = false;
		while(cellIterator.hasNext()){
			cell = cellIterator.next();
			try{
				if(cell.getNumericCellValue()==from){
					cell.setCellValue(to);
					toReturn=true;
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return toReturn;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceAll(String from, String to){
		from = from.replaceAll("[^\\x00-\\x7F]", "");
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		boolean toReturn = false;
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			Cell cell = null;
			while(cellIterator.hasNext()){
				cell = cellIterator.next();
				try{
					if(cell.getStringCellValue().replaceAll("[^\\x00-\\x7F]", "").compareTo(from)==0){
						cell.setCellValue(to);
						toReturn=true;
					}
				}catch(IllegalStateException e){
					//Do nothing. Just ignore this cell
				}
			}
		}
		return toReturn;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Boolean if the specified value was found or not.
	 */
	public boolean findAndReplaceAll(double from, double to){
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		boolean toReturn = false;
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			Cell cell = null;
			while(cellIterator.hasNext()){
				cell = cellIterator.next();
				try{
					if(cell.getNumericCellValue()==from){
						cell.setCellValue(to);
						toReturn=true;
					}
				}catch(IllegalStateException e){
					//Do nothing. Just ignore this cell
				}
			}
		}
		return toReturn;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Column index of the first found element. -1 if the element was not found
	 */
	public int findInRow(int row, String from){
		from = from.replaceAll("[^\\x00-\\x7F]", "");
		XSSFRow rowCells = this.currentSheet.getRow(row);
		Iterator<Cell> cellIterator = rowCells.cellIterator();
		Cell cell = null;
		while(cellIterator.hasNext()){
			cell = cellIterator.next();
			try{
				if(cell.getStringCellValue().replaceAll("[^\\x00-\\x7F]", "").compareTo(from)==0){
					return cell.getColumnIndex();
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return -1;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Column index of the first found element. -1 if the element was not found
	 */
	public int findInRow(int row, double from){
		XSSFRow rowCells = this.currentSheet.getRow(row);
		Iterator<Cell> cellIterator = rowCells.cellIterator();
		Cell cell = null;
		while(cellIterator.hasNext()){
			cell = cellIterator.next();
			try{
				if(cell.getNumericCellValue()==from){
					return cell.getColumnIndex();
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return -1;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Row index of the first found element. -1 if the element was not found
	 */
	public int findInColumn(int column, double from){
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			try{
				if(row.getCell(column).getNumericCellValue()==from){
					return row.getRowNum();
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return -1;
	}
	
	/**
	 * This function searches in the current sheet for any occurrence of the value 'from'
	 * and replaces it by the value 'to'
	 * @param from
	 * @param to
	 * @return Row index of the first found element. -1 if the element was not found
	 */
	public int findInColumn(int column, String from){
		from = from.replaceAll("[^\\x00-\\x7F]", "");
		Iterator<Row> rowIterator = this.currentSheet.iterator();
		Row row = null;
		while(rowIterator.hasNext()){
			row = rowIterator.next();
			try{
				if(row.getCell(column).getStringCellValue().replaceAll("[^\\x00-\\x7F]", "").compareTo(from)==0){
					return row.getRowNum();
				}
			}catch(IllegalStateException e){
				//Do nothing. Just ignore this cell
			}
		}
		return -1;
	}
	/**
	 * Paste a group of cell values(matrix) in the current sheet starting in the 
	 * (row,column) cell forward and down.
	 * @param row
	 * @param column
	 * @param matrix
	 * @return Boolean if any paste was performed.
	 */
	public boolean paste(int row, int column, Object matrixObject){
		JSONArray matrix = caller.checkJSONArray(matrixObject);
		if(matrix==null||matrix.length()==0)
			return false;
		
		for(int i=matrix.length()-1;i>=0;i--){
			try {
				JSONArray rowCells = matrix.getJSONArray(i);
				//Add a new row in the current sheet
				if(this.currentSheet.getRow(row+i)==null)
					this.currentSheet.createRow(row+i);
				
				if(rowCells.get(0) instanceof XSSFCell){
					XSSFCell cell = null;
					
					for(int j=rowCells.length()-1;j>=0;j--){
						
						if(this.currentSheet.getRow(row+i).getCell(column+j)==null)
							this.currentSheet.getRow(row+i).createCell(column+j);
						
						cell = (XSSFCell)(rowCells.get(j));
						int  type = cell.getCellType();
						switch(type){
						case Cell.CELL_TYPE_BOOLEAN:
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.getBooleanCellValue());;
							break;
						case Cell.CELL_TYPE_ERROR:
							this.currentSheet.getRow(row+i).getCell(column+j).setCellErrorValue(cell.getErrorCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							this.currentSheet.getRow(row+i).getCell(column+j).setCellFormula(cell.getCellFormula());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.getStringCellValue());
							break;
						}
						
					}
				}
				if(rowCells.get(0) instanceof JSONObject){
					JSONObject cell = null;
					for(int j=rowCells.length()-1;j>=0;j--){
						
						if(this.currentSheet.getRow(row+i).getCell(column+j)==null)
							this.currentSheet.getRow(row+i).createCell(column+j);
						
						cell = rowCells.getJSONObject(j);
						String  type = cell.optString("type", "");
						if(type.compareTo("formula")==0){
							this.currentSheet.getRow(row+i).getCell(column+j).setCellFormula(cell.optString("value"));
						}
						if(type.compareTo("numeric")==0){
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.optDouble("value"));
						}
						if(type.compareTo("string")==0){
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.optString("value"));
						}
						if(type.compareTo("boolean")==0){
							this.currentSheet.getRow(row+i).getCell(column+j).setCellValue(cell.optBoolean("value"));
						}
						if(type.compareTo("date")==0){
							//????
						}
					}
				}
			} catch (JSONException e) {
				caller.appendError("ExcelWorkbook::paste", e.toString());
				return false;
			}
		}
		return true;
	}
	
	/**
	 * Copy a group of cell values(matrix) from the current sheet starting in the 
	 * (row0,column0) cell and ending in the (row0,column0) cell.
	 * @param column
	 * @param row
	 * @param matrix
	 * @return Boolean if the specified value was found or not.
	 * @throws JSONException 
	 */
	public JSONArray copy(int row0, int column0, int row1, int column1, Object optionsObject) throws JSONException{
		JSONObject options = caller.checkParameter(optionsObject);
		int nRows = this.currentSheet. getLastRowNum(); 
		String val = options.optString("format","cell");
		boolean formulas = options.optBoolean("formulas",false);
		JSONArray toReturn = new JSONArray();
		if(row0>row1){
			int tmp=row0;
			row0=row1;
			row1=tmp;
		}
		
		if(column0>column1){
			int tmp=column0;
			column0=column1;
			column1=tmp;
		}
		XSSFCell tmp;
		JSONObject cell;
		if(val.compareTo("json")==0){
			for(int i=row0;i<=row1;i++){
				if(i<=nRows){
					JSONArray row = new JSONArray();
					for(int j=column0;j<=column1;j++){
						tmp = this.currentSheet.getRow(i).getCell(j);
						if(tmp==null)
							row.put(0);
						else{
							cell = new JSONObject();
							if(tmp.getCellType()==Cell.CELL_TYPE_FORMULA&&formulas){
								cell.append("type", "formula");
								cell.append("value", tmp.getCellFormula());
								row.put(cell);
							}
							else{
								switch(tmp.getCellType()){
									case Cell.CELL_TYPE_BOOLEAN:
										cell.append("type", "boolean");
										cell.append("value", tmp.getBooleanCellValue());
										row.put(cell);
										break;
									case Cell.CELL_TYPE_NUMERIC:
										cell.append("type", "numeric");
										cell.append("value", tmp.getNumericCellValue());
										row.put(cell);
										break;
									case Cell.CELL_TYPE_STRING:
										cell.append("type", "string");
										cell.append("value", tmp.getStringCellValue());
										row.put(cell);
										break;
									case Cell.CELL_TYPE_FORMULA:
										try{
											cell.append("value", tmp.getNumericCellValue());
											cell.append("type", "numeric");
										}
										catch(IllegalStateException e ){
											cell.append("value", tmp.getRawValue());
											cell.append("type", "string");
										}
										row.put(cell);
										break;
									default:
										row.put(0);
								}
							}
						}
					}
					toReturn.put(row);
				}
				
				
			}
		}
		else{
			if(val.compareTo("value")==0){
				for(int i=row0;i<=row1;i++){
					if(i<=nRows){
						JSONArray row = new JSONArray();
						for(int j=column0;j<=column1;j++){
							tmp = this.currentSheet.getRow(i).getCell(j);
							if(tmp==null)
								row.put(0);
							else{
								switch(tmp.getCellType()){
									case Cell.CELL_TYPE_BOOLEAN:
										row.put(tmp.getBooleanCellValue());
										break;
									case Cell.CELL_TYPE_NUMERIC:
										row.put(tmp.getNumericCellValue());
										break;
									case Cell.CELL_TYPE_STRING:
										row.put(tmp.getStringCellValue());
										break;
									case Cell.CELL_TYPE_FORMULA:
										try{
											row.put(tmp.getNumericCellValue());
										}
										catch(IllegalStateException e ){
											row.put(tmp.getRawValue());
										}
										break;
									default:
										row.put(0);
								}
							}
						}
						toReturn.put(row);
					}
				}
			}else{
				for(int i=row0;i<=row1;i++){
					JSONArray row = new JSONArray();
					for(int j=column0;j<=column1;j++){
						row.put(this.currentSheet.getRow(i).getCell(j));
					}
					toReturn.put(row);
				}
			}
		}
		return toReturn;
	}
	
	/**
	 * Copy a group of cell values(matrix) from the current sheet using the specified
	 * rows and columns index.
	 * @param rows
	 * @param columns
	 * @param options
	 * @throws JSONException 
	 */
	public JSONArray copy(int[] rows, int[] columns, Object optionsObject) throws JSONException{
		//System.out.println("Aqui "+rows.length);
		JSONObject options = caller.checkParameter(optionsObject);
		String val = options.optString("format","cell");
		boolean formulas = options.optBoolean("formulas",false);
		JSONArray toReturn = new JSONArray();
		int nRows = this.currentSheet.getLastRowNum(); 
		XSSFCell tmp;
		JSONObject cell;
		if(val.compareTo("json")==0){
			for(int i : rows){
				if(i<=nRows){
					JSONArray row = new JSONArray();
					for(int j:columns){
							tmp = this.currentSheet.getRow(i).getCell(j);
							if(tmp==null)
								row.put(0);
							else{
								cell = new JSONObject();
								if(tmp.getCellType()==Cell.CELL_TYPE_FORMULA&&formulas){
									cell.append("type", "formula");
									cell.append("value", tmp.getCellFormula());
									row.put(cell);
								}
								else{
									switch(tmp.getCellType()){
										case Cell.CELL_TYPE_BOOLEAN:
											cell.append("type", "boolean");
											cell.append("value", tmp.getBooleanCellValue());
											row.put(cell);
											break;
										case Cell.CELL_TYPE_NUMERIC:
											cell.append("type", "numeric");
											cell.append("value", tmp.getNumericCellValue());
											row.put(cell);
											break;
										case Cell.CELL_TYPE_STRING:
											cell.append("type", "string");
											cell.append("value", tmp.getStringCellValue());
											row.put(cell);
											break;
										case Cell.CELL_TYPE_FORMULA:
											try{
												cell.append("value", tmp.getNumericCellValue());
												cell.append("type", "numeric");
											}
											catch(IllegalStateException e ){
												cell.append("value", tmp.getRawValue());
												cell.append("type", "string");
											}
											row.put(cell);
											break;
										default:
											row.put(0);
									}
								}
							}
					}
					toReturn.put(row);
				}
			}
		}
		else{
			if(val.compareTo("value")==0){
				for(int i:rows){
					if(i<=nRows){
						//System.out.println(i);
						JSONArray row = new JSONArray();
						for(int j:columns){
							tmp = this.currentSheet.getRow(i).getCell(j);
							if(tmp==null)
								row.put(0);
							else{
								switch(tmp.getCellType()){
									case Cell.CELL_TYPE_BOOLEAN:
										row.put(tmp.getBooleanCellValue());
										break;
									case Cell.CELL_TYPE_NUMERIC:
										row.put(tmp.getNumericCellValue());
										break;
									case Cell.CELL_TYPE_STRING:
										row.put(tmp.getStringCellValue());
										break;
									case Cell.CELL_TYPE_FORMULA:
										try{
											row.put(tmp.getNumericCellValue());
										}
										catch(IllegalStateException e ){
											row.put(tmp.getRawValue());
										}
										break;
									default:
										row.put(0);
								}
							}
						}
						toReturn.put(row);
					}	
				}
			}else{
				for(int i:rows){
					JSONArray row = new JSONArray();
					for(int j:columns){
						row.put(this.currentSheet.getRow(i).getCell(j));
					}
					toReturn.put(row);
				}
			}
		}
		
		return toReturn;
	}
	
	public void put(int row, int col, Object cellObject){
		if(this.currentSheet.getRow(row)==null){
			this.currentSheet.createRow(row);
		}
		if(this.currentSheet.getRow(row).getCell(col)==null){
			this.currentSheet.getRow(row).createCell(col);
		}
		JSONObject cell = caller.checkParameter(cellObject);
		if(cell!=null){
			String type = cell.optString("type","string");
			if(type.compareTo("formula")==0){
				this.currentSheet.getRow(row).getCell(col).setCellFormula(cell.optString("value",""));
			}
			else{
				if(type.compareTo("string")==0){
					this.currentSheet.getRow(row).getCell(col).setCellValue(cell.optString("value",""));
				}
				else{
					if(type.compareTo("numeric")==0){
						this.currentSheet.getRow(row).getCell(col).setCellValue(cell.optDouble("value",0));
					}
				}
			}
		}
	}
	
	public int nRows(int column, int startIndex){
		if(startIndex<0)
			startIndex=0;
		return this.currentSheet.getLastRowNum();
	}
}
