/**
 * 
 */
package com.jacob.excel.sheet.items;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.excel.Excel;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.script.RunCommand;
import com.jacob.excel.sheet.Sheet;


/**
 * @author Mircea Sirghi
 *
 */
public class Cell extends ExcelBase implements ICom {
	Dispatch cell;
	
	public Cell(Cells cells, int i)
	{
		cell = Dispatch.invoke(cells.getActiveX(), "Item", Dispatch.Get, new Object[] { i }, new int[0]).getDispatch();
	}
	
	public Cell(Sheet sheet, int i, int j)
	{
		cell = Dispatch.invoke(sheet.getActiveX(), "Cells", Dispatch.Get, new Object[] { i, j }, new int[0]).getDispatch();
	}
	
	public int getRow()
	{	
		return getIntField(cell, "Row");
	}
	
	public int getColumn()
	{	
		return getIntField(cell, "Column");
	}
	
	public String getValue()
	{
		
		return getStringField(cell, "Value");
	}
	
	public String getFormula()
	{
		return getStringField(cell, "Formula");
	}
	
	public Double getDate()
	{		
		return getDateField(cell, "Formula");
	}
	
	public String getError()
	{
		return getStringField(cell, "Value");
	}
	
	
	public int getPivotType()
	{
		Dispatch cellPivot = Dispatch.invoke(cell, "PivotCell", Dispatch.Get, new Object[] {  }, new int[1]).getDispatch();
		
		int cellPivotType = Dispatch.get(cellPivot, "PivotCellType").toInt();	
		
		return cellPivotType;
	}
	
	public CellType getCellType(Excel excel)
	{
		
		CellType cellType = CellType.None;
		try
		{
			boolean cellTypeText = Dispatch.invoke(excel.getActiveX(), "IsText", Dispatch.Get, new Object[] { cell }, new int[1]).getBoolean();		
			boolean cellTypeLogical = Dispatch.invoke(excel.getActiveX(), "IsLogical", Dispatch.Get, new Object[] { cell }, new int[1]).getBoolean();
			boolean cellTypeErr = Dispatch.invoke(excel.getActiveX(), "IsErr", Dispatch.Get, new Object[] { cell }, new int[1]).getBoolean();
			//boolean cellTypeDate = Dispatch.invoke(excel.getActiveX(), "IsDate", Dispatch.Get, new Object[] { cell }, new int[1]).getBoolean();
			boolean cellTypeDate = RunCommand.run("IsDate", new Object [] {cell}).equalsIgnoreCase("true");
			boolean cellTypeBlank = RunCommand.run("IsEmpty", new Object [] {cell}).equalsIgnoreCase("true");
			boolean cellTypeNumeric = RunCommand.run("IsNumeric", new Object [] {cell}).equalsIgnoreCase("true");
			
			if(cellTypeText) cellType = CellType.Text;
			else if(cellTypeLogical) cellType = CellType.Logical;
			else if(cellTypeErr) cellType = CellType.Error;
			else if(cellTypeDate) cellType = CellType.Date;
			else if(cellTypeBlank) cellType = CellType.Blank;
			else if(cellTypeNumeric) cellType = CellType.Numeric;
		}
		catch(Exception e)
		{
			e.printStackTrace();			
		}
		
////		Case IsEmpty(c): CellType = "Blank"
////		Case Application.IsText(c): CellType = "Text"
////		Case Application.IsLogical(c): CellType = "Logical"
////		Case Application.IsErr(c): CellType = "Error"
////		Case Application.IsDate(c): CellType = "Date"
////		Case InStr(1, c.Text, ":") <> 0: CellType = "Time"
////		Case IsNumeric(c): CellType = "Value"
		
		return cellType;
	}
	
	public CellType getCellType(Dispatch excel)
	{		
		CellType cellType = RunCommand.getCellType(cell, excel);
		
		return cellType;
	}
	
	public Dispatch getActiveX() {		
		return cell;
	}
	
	@Override
	public void Close() {
		// TODO Auto-generated method stub
		
	}
	
	public String readRealValue(Excel excel)
	{
		//Cell cell = sheet.getCell(i,j);
		
		CellType cellType = getCellType(excel.getActiveX());
    	
    	String val = null;
    	            	            	
    	switch(cellType)
    	{
    	case None:
    		break; 
    	case Text: 
    		val = getValue();
    		break;
    	case Logical:
    		val = getFormula();
    	case Date:            		               	
    		try
    		{	                 	
	    		Integer dateToInt = Integer.parseInt(getFormula());
	    		if(dateToInt != null)
	    		{
	    			 val = ExcelDateParse(dateToInt).toString();			    			 
	    		}
    		}
    		catch(Exception e)            		
    		{
    			val = null;
    		}
    		break;
    	case Blank:
    		val = "";
    		break;
    	case Error: 
    		val = getError();
    	case Numeric:
    		val = getFormula();            		
    		break; 
    	case Value:
    		val = getFormula();            		
    		break; 
    	}
    	
    	return val;
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
