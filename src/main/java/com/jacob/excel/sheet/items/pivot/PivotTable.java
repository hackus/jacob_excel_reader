/**
 * 
 */
package com.jacob.excel.sheet.items.pivot;

import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.sheet.Sheet;
import com.jacob.excel.sheet.items.Cells;
import com.jacob.excel.sheet.items.pivot.PivotRange.PivotRangeType;


/**
 * @author Mircea Sirghi
 *
 */
public class PivotTable extends ExcelBase implements ICom {
	private Dispatch pt;
	
	public PivotTable(Sheet sheet, String name)
	{
		pt = Dispatch.invoke(
				sheet.getActiveX(),
				"PivotTables",
				Dispatch.Get,
				new Object[] { name },
				new int[1]
			 ).toDispatch();
	}

	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return pt;
	}
	
	public void refresh()
	{
		Dispatch.call(pt, "RefreshTable");
	}
	
	public PivotRange getRange()
	{
		//Dispatch range = Dispatch.invoke(pt, "TableRange1", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
		
		return new PivotRange(this);
	}
	
	public PivotRange getRowRange()
	{
		//Dispatch range = Dispatch.invoke(pt, "TableRange1", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
		
		return new PivotRange(this, PivotRangeType.RowRange);
	}
	
	public PivotRange getColumnRange()
	{
		return new PivotRange(this, PivotRangeType.ColumnRange);
	}
	
	public Cells getCells()
	{
		//Dispatch cells = Dispatch.invoke(getRange(), "Cells", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
		
		return new Cells(getRange());
	}
	
	public PivotRows getRows()
	{
		return new PivotRows(this);
	}
	
	public PivotColumns getColumns()
	{
		return new PivotColumns(this);
	}

	@Override
	public void Close() {
		// TODO Auto-generated method stub
		
	}
}
