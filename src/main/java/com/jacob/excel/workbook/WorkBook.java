/**
 * 
 */
package com.jacob.excel.workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.sheet.Sheets;

/**
 * @author Mircea Sirghi
 *
 */
public class WorkBook extends ExcelBase implements ICom {
	//private WorkBooks workBooks;
	Dispatch workBook;
	
	List<Sheets> sheetsList = new ArrayList<Sheets>();
	
	public WorkBook(WorkBooks workBooks, File file)
	{
		Wait(1);
		//this.workBooks = workBooks;	
		workBook = Dispatch.call(workBooks.getActiveX(), "Open", file.getAbsolutePath()).toDispatch();
	}
	
	public Sheets createSheets()
	{
		Sheets sheetsObj = new Sheets(this);
		sheetsList.add(sheetsObj);
		return sheetsObj;
	}
	
	public List<Sheets> getSheets()
	{	
		return sheetsList;
	}

	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return workBook;
	}
	
	public void Save()
	{
		 Dispatch.call(workBook, "Save");
	}
	
	public void Close()
	{
	   for(Sheets sheetsObj : sheetsList)
	   {
		   sheetsObj.Close();
	   }
	   com.jacob.com.Variant f = new com.jacob.com.Variant(true);
       Dispatch.call(workBook, "Close", f);
	}
}
