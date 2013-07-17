/**
 * 
 */
package com.jacob.excel.sheet;

import java.util.ArrayList;
import java.util.List;

import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.workbook.WorkBook;


/**
 * @author Mircea Sirghi
 *
 */
public class Sheets extends ExcelBase implements ICom  {
	private Dispatch sheets;	
	
	List<Sheet> sheetList = new ArrayList<Sheet>();
	
	public Sheets(WorkBook workBook)
	{
		Wait(1);
		sheets = Dispatch.get(workBook.getActiveX(), "Sheets").toDispatch();
	}

	public Sheet getSheet(String sheetName)
	{
		Sheet sheet = new Sheet(this, sheetName);
		sheetList.add(sheet);
		return new Sheet(this, sheetName);
	}
	
	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return sheets;
	}

	@Override
	public void Close() {
		for(Sheet sheet : sheetList)
		{
			sheet.Close();
		}
	}	
}
