package com.jacob.excel;
/**
 * 
 */

/**
 * @author Mircea Sirghi
 */
import java.io.File;
import java.io.FileInputStream;
//import this if desired to handle an IOException separately
import java.io.IOException;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
//import this if desired to handle an InvalidFormatException separately
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.jacob.excel.reader.ReadPivotTableByRows;
import com.jacob.excel.script.RunCommand;

public class ReadExcel {

//	static {
//	    System.loadLibrary("jacob-1.17-M2-x64.dll");
//	}
	
	/**
	 * @param args is not used.  
	 */
	public static void main(String[] args) {
		
		try {
			ReadPivotTableByRows.read(new File("lib\\amazon_referrals_2009.xlsx"), "Pivot Table 1");
		}
		catch (Exception e) {
			e.printStackTrace();			
		}
	}
}
