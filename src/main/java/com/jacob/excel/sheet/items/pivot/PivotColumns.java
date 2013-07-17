/**
 * 
 */
package com.jacob.excel.sheet.items.pivot;


import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.base.IComList;

/**
 * @author Mircea Sirghi
 *
 */
public class PivotColumns extends ExcelBase implements IComList {
	private Dispatch cols;
	
	PivotTable pt;
	
	public PivotColumns(PivotTable pt)
	{	
		this.pt = pt;
		cols = Dispatch.invoke(pt.getRange().getActiveX(), "Columns", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
	}
	
	@Override
	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return cols;
	}
	
	

	@Override
	public void Close() {
		
		
	}

	@Override
	public int Count() {		
		int number = Dispatch.get(cols, "Count").getInt();
		
		//int number = Dispatch.get(pt.getColumnRange().getActiveX(), "Count").getInt();
		
		return number;
	}
}
