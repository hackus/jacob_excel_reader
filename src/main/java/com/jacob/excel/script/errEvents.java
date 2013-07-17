/**
 * 
 */
package com.jacob.excel.script;

import com.jacob.com.Variant;

/**
 * @author Mircea Sirghi
 *
 */
public class errEvents {
	public void Error(Variant[] args)
    {
        System.out.println("java callback for error!");
    }
    public void Timeout(Variant[] args)
    {
        System.out.println("java callback for timeout!");
    }
}
