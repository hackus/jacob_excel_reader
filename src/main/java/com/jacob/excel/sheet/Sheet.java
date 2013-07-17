/**
 * 
 */
package com.jacob.excel.sheet;

import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.sheet.items.Cell;
import com.jacob.excel.sheet.items.pivot.PivotTable;


/**
 * @author Mircea Sirghi
 *
 */
public class Sheet extends ExcelBase implements ICom  {
	Dispatch sheet;
	
	public Sheet(Sheets sheets, String name)
	{
		sheet = Dispatch.invoke(sheets.getActiveX(), "Item", Dispatch.Get, new Object[] { name }, new int[0]).getDispatch();
	}
	
	public PivotTable getPivotTable(String name)
	{
		return new PivotTable(this, name);
	}

	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return sheet;
	}

	public Cell getCell(int i, int j)
	{
		return new Cell(this, i, j);
	}
	
	@Override
	public void Close() {
		sheet.safeRelease();		
	}
}
