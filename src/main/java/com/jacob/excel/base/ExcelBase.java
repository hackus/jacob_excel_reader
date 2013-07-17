/**
 * 
 */
package com.jacob.excel.base;

import java.io.File;

import com.jacob.com.Dispatch;

/**
 * @author Mircea Sirghi
 *
 */
public abstract class ExcelBase {
	public File file;
	
	public Integer getIntField(Dispatch obj, String fieldName)
	{
		Integer val=null;
		try
		{
			val =  Dispatch.get(obj, fieldName).getInt();
			//System.out.print(fieldName + ":["+ val+"]");
		}
		catch(Exception e)
		{
			//System.out.print(fieldName + ":[nothing]");
		}
		return val;
	}
	
	public String getStringField(Dispatch obj, String fieldName)
	{
		String val=null;
		try
		{
			val = Dispatch.get(obj, fieldName).getString();
			//System.out.print(fieldName + ":["+ val+"]");
		}
		catch(Exception e)
		{
			//System.out.print(fieldName + ":[nothing]");
		}
		return val;
	}
	
	public Double getDateField(Dispatch obj, String fieldName)
	{
		Double val=null;
		try
		{
			val = Dispatch.get(obj, fieldName).getDate();
			//System.out.print(fieldName + ":["+ val+"]");
		}
		catch(Exception e)
		{
			//System.out.print(fieldName + ":[nothing]");
		}
		return val;
	}
	
	protected void Wait(double timeout)
	{
		timeout = timeout * 1000;
		try {
			Thread.sleep((int)timeout);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
