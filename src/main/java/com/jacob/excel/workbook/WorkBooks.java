/**
 * 
 */
package com.jacob.excel.workbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


import com.jacob.com.Dispatch;
import com.jacob.excel.Excel;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;

/**
 * @author Mircea Sirghi
 *
 */
public class WorkBooks extends ExcelBase implements ICom {

	private Dispatch workbooks;
	
	List<WorkBook> workBooksList = new ArrayList<WorkBook>();
	
	//Excel excel;
	
	public WorkBooks(Excel excel)
	{
		Wait(1);
		//this.excel = excel;
		workbooks = excel.getActiveX().getProperty("Workbooks").getDispatch();		
		Dispatch.call(workbooks, "Add");
	}
	
	public int Count()
	{	
		int id = Dispatch.get(workbooks, "Count").getInt();
		return id;
	}
	
	public Dispatch getActiveX()
	{
		return workbooks;
	}
	
	public WorkBook createWorkBook(File file)
	{
		WorkBook wb = new WorkBook(this, file);
		workBooksList.add(wb);
		return wb;
	}
	
	public List<WorkBook> getWorkBooks()
	{		
		return workBooksList;
	}

	@Override
	public void Close() {
		for(WorkBook wb : workBooksList)
		{
			wb.Close();
		}
		workbooks.safeRelease();
	}
}
