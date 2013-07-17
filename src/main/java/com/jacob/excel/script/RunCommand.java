/**
 * 
 */
package com.jacob.excel.script;


import com.jacob.com.*;
import com.jacob.excel.sheet.items.CellType;

/**
 * @author Mircea Sirghi
 *
 */
public class RunCommand {
	 public static String run(String args[])
	 {
	     System.runFinalizersOnExit(true);
	     String lang = "VBScript";
	     //Dispatch sControl = new Dispatch("ScriptControl");
	     //Dispatch sControl = new Dispatch("IWBScriptControl");	     
	     //Dispatch sControl = new Dispatch("MSScriptControl.ScriptControl.1");
	     Dispatch sControl = new Dispatch("ScriptControl");
	     Dispatch.put(sControl, "Language", lang);
 	     Dispatch.put(sControl, "AllowUI", new Variant(true));
	     // instantiate an event target object
  	     errEvents te = new errEvents();
	     // hook it up to the sControl source
	     DispatchEvents de = new DispatchEvents(sControl, te);
	     // run an expression from the command line
	     
	     String str = Dispatch.call(sControl, "Eval", args[0]).toString();
	     System.out.println("eval("+args[0]+")=["+str +"]");
	     
	     return str;
	  }	
	 
	 public static String run(String functionName, Object args[])
	 {
	     System.runFinalizersOnExit(true);
	     String lang = "VBScript";
	     //Dispatch sControl = new Dispatch("ScriptControl");
	     //Dispatch sControl = new Dispatch("IWBScriptControl");	     
	     //Dispatch sControl = new Dispatch("MSScriptControl.ScriptControl.1");
	     Dispatch sControl = new Dispatch("ScriptControl");
	     Dispatch.put(sControl, "Language", lang);
 	     Dispatch.put(sControl, "AllowUI", new Variant(true));
	     // instantiate an event target object
  	     errEvents te = new errEvents();
	     // hook it up to the sControl source
	     DispatchEvents de = new DispatchEvents(sControl, te);
	     // run an expression from the command line
	     
	     //Dispatch.call(sControl, "Execute", args[0], args[1]).toString();
	     
//	     Declare(scr; s)
//	     Object(scr; "MSScriptControl.ScriptControl")
//	     scr =  CreateObject("MSScriptControl.ScriptControl")
//	     s = "Function T: T=IsEmpty() : End Function"
//	     scr.Language = "VBScript"
//	     scr.AddCode (s)
//	     scr.Run ("T")	     
	     
	     String objectsList = "";
	     
	     for(int i=0;i<args.length;i++)
	     {
	    	 Dispatch.call(sControl, "AddObject", new Object [] {"test" + i, args[i]}).toString();
	    	 objectsList += "test" +i;
	    	 
	    	 if((i+1)<args.length)
	    		 objectsList += ",";
	     }	     
	     
	     //String script =  "Function T: T="+functionName+"(test) : End Function";
	     String script =  "Function T: T=" +functionName+ "(" +objectsList+ ") : End Function";
	     
	     Dispatch.call(sControl, "AddCode", new Object [] {script}).toString();
	     
	     //Dispatch.call(sControl, "AddObject", new Object [] {"test", args[1]}).toString();
	     
	     //String str = Dispatch.invoke(sControl, "Run" , Dispatch.Get, new Object [] {"T"}, new int[1]).toString();
	     
	     //String str = Dispatch.call(sControl, "ExecuteStatement", new Object [] {"IsEmpty(test)"}).toString();
	     
	     String str = Dispatch.call(sControl, "Run", new Object [] {"T"}).toString();
	     //System.out.println("");
	     
	     return str;
	  }	
	 
	 public static CellType getCellType(Dispatch cell, Dispatch excel)
	 {
		 System.runFinalizersOnExit(true);
	     String lang = "VBScript";	    
	     Dispatch sControl = new Dispatch("ScriptControl");
	     Dispatch.put(sControl, "Language", lang);
		 Dispatch.put(sControl, "AllowUI", new Variant(true));	    
		 errEvents te = new errEvents();	    
	     DispatchEvents de = new DispatchEvents(sControl, te);
	     // run an expression from the command line
	     
	     //Dispatch.call(sControl, "Execute", args[0], args[1]).toString();
	     
	//     Declare(scr; s)
	//     Object(scr; "MSScriptControl.ScriptControl")
	//     scr =  CreateObject("MSScriptControl.ScriptControl")
	//     s = "Function T: T=IsEmpty() : End Function"
	//     scr.Language = "VBScript"
	//     scr.AddCode (s)
	//     scr.Run ("T")	   
	     
	    String script = "Function CellType(c) : "
	    + "    Application.Volatile" + " : "
	    //+ "    msgbox(c)" + " : "
	    + "    Set c = c.Range(\"A1\")" + " : "
	    + "    Select Case True" + " : "
	    + "        Case IsEmpty(c): CellType = \"Blank\"" + " : "
	    + "        Case Application.IsText(c): CellType = \"Text\"" + " : "
	    + "        Case Application.IsLogical(c): CellType = \"Logical\"" + " : "
	    + "        Case Application.IsErr(c): CellType = \"Error\"" + " : "
	    + "        Case IsDate(c): CellType = \"Date\"" + " : "
	    //+ "        Case InStr(1, c.Text, \":\") <> 0: CellType = \"Time\"" + " : "
	    + "        Case IsNumeric(c): CellType = \"Value\"" + " : "
	    + "    End Select" + " : "
	    //+ "    msgbox(CellType)" + " : "
	    + " : End Function";
	     
	    
	    Dispatch.call(sControl, "AddObject", new Object [] {"Application", excel}).toString();	  
	    
	    Dispatch.call(sControl, "AddObject", new Object [] {"testCell", cell}).toString();
	     
	    Dispatch.call(sControl, "AddCode", new Object [] {script}).toString();
	     
	     
	    String str = Dispatch.call(sControl, "Eval", new Object [] {"CellType(testCell)"}).toString();
	    
	    //String str = Dispatch.call(sControl, "Run", new Object [] {"CellType" , "test"}).toString();
	    
	    return CellType.fromString(str);
	 	
	 }
}
