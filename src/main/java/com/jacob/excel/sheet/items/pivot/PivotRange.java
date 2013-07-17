/**
 * 
 */
package com.jacob.excel.sheet.items.pivot;

import com.jacob.com.Dispatch;
import com.jacob.excel.base.ExcelBase;
import com.jacob.excel.base.ICom;
import com.jacob.excel.sheet.items.CellType;


/**
 * @author Mircea Sirghi
 *
 */
public class PivotRange extends ExcelBase implements ICom {
	Dispatch range; 
	
	public enum PivotRangeType
	{
		RowRange("RowRange"),
		ColumnRange("ColumnRange"),
		TableRange1("TableRange1");
		
		private String text;

		PivotRangeType(String text) {
			this.text = text;
		}

		public String getText() {
			return this.text;
		}

		public static PivotRangeType fromString(String text) {
			if (text != null) {
			  for (PivotRangeType b : PivotRangeType.values()) {
				if (text.equalsIgnoreCase(b.text)) {
				  return b;
				}
			  }
			}
			return null;
		}
	}
	
	public PivotRange(PivotTable pt)
	{
		range = Dispatch.invoke(pt.getActiveX(), "TableRange1", Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
	}
	
	public PivotRange(PivotTable pt, PivotRangeType rangeType)
	{
		range = Dispatch.invoke(pt.getActiveX(), rangeType.name(), Dispatch.Get, new Object[] {  }, new int[1]).toDispatch();
	}
	
	public Dispatch getActiveX() {
		// TODO Auto-generated method stub
		return range;
	}

	@Override
	public void Close() {
		// TODO Auto-generated method stub
		
	}
}
