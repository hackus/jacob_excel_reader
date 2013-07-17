/**
 * 
 */
package com.jacob.excel.sheet.items;


import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.sheet.items.pivot.PivotRange;

/**
 * @author Mircea Sirghi
 *
 */
public class Cells extends ExcelBase implements ICom  {

	Dispatch cells;
	
	public Cells(PivotRange range)
	{
		cells = Dispatch.invoke(range.getActiveX(), "Cells", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
	}
	
	public Cell getCell(int i)
	{
		return new Cell(this, i);
	}
	
	public int Count()
	{
		int cellsNumber = Dispatch.get(cells, "Count").getInt();
		return cellsNumber;
	}

	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return cells;
	}

	@Override
	public void Close() {
		// TODO Auto-generated method stub
		
	}
}
