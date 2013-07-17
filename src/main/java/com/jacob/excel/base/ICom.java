/**
 * 
 */
package com.jacob.excel.base;

import com.jacob.com.Dispatch;

/**
 * @author Mircea Sirghi
 *
 */
public interface ICom {
	public Dispatch getActiveX();
	
	public void Close();
}
