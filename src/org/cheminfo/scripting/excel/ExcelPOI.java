package org.cheminfo.scripting.excel;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.cheminfo.function.Function;
import org.cheminfo.function.scripting.SecureFileManager;

/**
 *  
 * @author acastillo 
 * 
 */
public class ExcelPOI extends Function{
	/**
	 * This function load an Excel work book. Supported formats: XLSX, XLS
	 * @param basedir
	 * @param key
	 * @param filename
	 * @return and extended XLSXWorkook or XLSWorkook object
	 */
	public ExcelWorkbook load(String basedir, String basedirkey, String filename){
		// If it is a URL we wont check security
		if (filename.trim().matches("^https?://.*$")) {
			try {
				return new ExcelWorkbook(basedir, basedirkey, filename, this);
			} catch (FileNotFoundException e) {
				this.appendError("ExcelWorkbook::load", "File "+filename+" not found");
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} else {
			String fullFilename = SecureFileManager.getValidatedFilename(
					basedir, basedirkey, filename);
			if (fullFilename == null)
				return null;
			try {
				return new ExcelWorkbook(basedir, basedirkey, fullFilename, this);
			} catch (IOException e) {
				this.appendError("ExcelWorkbook::load", "IOException while reading "+filename);
				e.printStackTrace();
			}
		}
		return null;
	}
	
	/**
	 * This function load an Excel work book. Supported formats: XLSX, XLS
	 * @param basedir
	 * @param key
	 * @param filename
	 * @return and extended XLSXWorkook or XLSWorkook object
	 */
	public ExcelWorkbook create(String basedir, String basedirkey, Object data){
		return null;
	}
}
