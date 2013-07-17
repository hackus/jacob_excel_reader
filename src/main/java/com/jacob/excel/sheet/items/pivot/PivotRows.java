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
public class PivotRows extends ExcelBase implements IComList {
	private Dispatch rows;
	
	PivotTable pt;
	
	public PivotRows(PivotTable pt)
	{	
		this.pt = pt;
		rows = Dispatch.invoke(pt.getRange().getActiveX(), "Rows", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
	}
	
	@Override
	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return rows;
	}
	
	@Override
	public void Close() {
		// TODO Auto-generated method stub
		
	}

	@Override
	public int Count() {		
		int number = Dispatch.get(rows, "Count").getInt();
		//int number = Dispatch.get(pt.getRowRange().getActiveX(), "Count").getInt();
		
		return number;
	}

}
