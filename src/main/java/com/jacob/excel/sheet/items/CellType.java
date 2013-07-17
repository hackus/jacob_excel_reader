/**
 * 
 */
package com.jacob.excel.sheet.items;

/**
 * @author Mircea Sirghi
 *
 */
public enum CellType {
	None("None"),
	Blank("Blank"), 
	Text("Text"),
	Value("Value"),
	Logical("Logical"),
	Error("Error"), 
	Date("Date"),
	Numeric("Numeric");

	private String text;

	CellType(String text) {
		this.text = text;
	}

	public String getText() {
		return this.text;
	}

	public static CellType fromString(String text) {
		if (text != null) {
		  for (CellType b : CellType.values()) {
			if (text.equalsIgnoreCase(b.text)) {
			  return b;
			}
		  }
		}
		return null;
	}
}
