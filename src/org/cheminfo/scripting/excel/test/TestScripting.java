package org.cheminfo.scripting.excel.test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

import junit.framework.Assert;

import org.cheminfo.function.scripting.ScriptingInstance;
import org.json.JSONObject;
import org.junit.Test;

public class TestScripting {



	/**
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		ScriptingInstance interpreter = new ScriptingInstance("");
		interpreter.setSafePath("./");
		String script = getContents(new File("tests/testFile.js"));
//		String script = getContents(new File("tests/testFile.js"));
//		String script = getContents(new File("tests/excelBasic.js"));
//		script = getContents(new File("tests/transparent.js"));
		JSONObject result = interpreter.runScript(script);
		System.out.println(result);

	}

	static public String getContents(File aFile) {
		try {
			System.out.println(aFile.getCanonicalPath());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// ...checks on aFile are elided
		StringBuilder contents = new StringBuilder();

		try {
			// use buffering, reading one line at a time
			// FileReader always assumes default encoding is OK!
			BufferedReader input = new BufferedReader(new FileReader(aFile));
			try {
				String line = null; // not declared within while loop
				/*
				 * readLine is a bit quirky : it returns the content of a line
				 * MINUS the newline. it returns null only for the END of the
				 * stream. it returns an empty String if two newlines appear in
				 * a row.
				 */
				while ((line = input.readLine()) != null) {
					contents.append(line);
					contents.append(System.getProperty("line.separator"));
				}
			} finally {
				input.close();
			}
		} catch (IOException ex) {
			ex.printStackTrace();
		}
		return contents.toString();
	}

}