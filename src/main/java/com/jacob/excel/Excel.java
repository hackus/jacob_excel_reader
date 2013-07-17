/**
 * 
 */
package com.jacob.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Variant;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.workbook.WorkBooks;

/**
 * @author Mircea Sirghi
 *
 */
public class Excel extends ExcelBase implements ICom{
	private ActiveXComponent excel;
	
	List<WorkBooks> wbs = new ArrayList<WorkBooks>();
	
	public Excel(boolean visible, File file)
	{		
		ComThread.InitSTA();

		excel = new ActiveXComponent("Excel.Application");
		Wait(1);
		  // This will open the excel if the property is set to true
		excel.setProperty("Visible", new Variant(visible));
		
		this.file = file;
	}
	
	public ActiveXComponent getActiveX()	
	{
		return excel;
	}
	
	public WorkBooks createWorkBooks()
	{
		WorkBooks wbsObject = new WorkBooks(this);
		wbs.add(wbsObject);
		return wbsObject;
	}

	@Override
	public void Close() {
		for(WorkBooks wbsObj : wbs)
		{
			wbsObj.Close();			
		}		
		excel.invoke("Quit", new Variant[0]);
		excel.safeRelease();
        ComThread.Release();
	}
}
