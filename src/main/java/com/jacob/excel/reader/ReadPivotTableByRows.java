/**
 * 
 */
package com.jacob.excel.reader;

import java.io.File;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;


import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.excel.Excel;
import com.jacob.excel.sheet.Sheet;
import com.jacob.excel.sheet.Sheets;
import com.jacob.excel.sheet.items.Cell;
import com.jacob.excel.sheet.items.CellType;
import com.jacob.excel.sheet.items.Cells;
import com.jacob.excel.sheet.items.pivot.PivotColumns;
import com.jacob.excel.sheet.items.pivot.PivotRows;
import com.jacob.excel.sheet.items.pivot.PivotTable;
import com.jacob.excel.workbook.WorkBook;
import com.jacob.excel.workbook.WorkBooks;

/**
 * @author Mircea Sirghi
 *
 */
public class ReadPivotTableByRows {
	public static void read(File file, String pivotSheetName) {
	    Excel excel = new Excel(false, file);

	    try {
	        WorkBooks workBooks = excel.createWorkBooks();	        
	        WorkBook workBook = workBooks.createWorkBook(file);
	        Sheets sheets =  workBook.createSheets();
	        Sheet sheet = sheets.getSheet(pivotSheetName);
	        PivotTable pt = sheet.getPivotTable("PivotTable1");
	        pt.refresh();
	        
	        
	        Cells cells = pt.getCells();
	        PivotRows rows = pt.getRows();
	        PivotColumns cols = pt.getColumns();

	        int rowNumber = rows.Count() + cells.getCell(1).getRow();
	        int colNumber = cols.Count() + cells.getCell(1).getColumn();
	        
	        for(int i=cells.getCell(1).getRow();i<rowNumber;i++)
	        {
	        	for(int j=cells.getCell(1).getColumn();j<colNumber;j++)
	        	{
	        		String val = sheet.getCell(i,j).readRealValue(excel);
	            	
	            	printStringValue(val);
	        	}
	        	System.out.println("");
	        }  
	    } catch (Exception e) {
	        e.printStackTrace();
	    } finally {
	    	excel.Close();	    	
	    }
	}
	
	private static void printStringValue(String value)
	{
		if(value != null)
			System.out.print("|" + value + "|");
		else 
			System.out.print("||");
		
	}
	private static void printDoubleValue(Double value)
	{
		if(value != null)
			System.out.print("|" + Double.toString(value) + "|");
		else 
			System.out.print("||");
		
	}
	private static void printIntValue(Integer value)
	{
		if(value != null)
			System.out.print("|" + Integer.toString(value) + "|");
		else 
			System.out.print("||");
		
	}
	
	public static Date ExcelDateParse(Integer ExcelDate){
		Date result = null;
		if(ExcelDate != null)
		{   
		    try{
		        GregorianCalendar gc = new GregorianCalendar(1900, Calendar.JANUARY, 1);
		        gc.add(Calendar.DATE, ExcelDate - 2);
		        result = gc.getTime();
		    } catch(RuntimeException e1) {}
		}
	    return result;
	}     
}
