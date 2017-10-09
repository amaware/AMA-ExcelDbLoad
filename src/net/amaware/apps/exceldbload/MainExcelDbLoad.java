package net.amaware.apps.exceldbload;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Vector;

//import net.amaware.serv.SourceServProperty;
import net.amaware.app.MainAppDataStore;
import net.amaware.autil.*;
import net.amaware.serv.HtmlTargetServ;

//
/**
 * @author AMAWARE - Angelo M Adduci
 * 
 */

public class MainExcelDbLoad {
	// set Properties file key names to being used
	//Properties file 
    final static String propFileName   = "MainExcelDbLoad.properties";
	//Architecture Common communication Class 
	static ACommDbFile acomm;
	//Architecture Framework Class
    //Target Services
	static HtmlTargetServ aHtmlServReport;
    //	
	static MainAppDataStore mainApp;
	//Application Classes
    //
	static String htmlTargetLine= "font-size:1.2em;color:blue;";
	static String htmlTargetLineWithBorder= htmlTargetLine+"color:orange;border:solid orange .1em;";
	//
        //
		public static void main(String[] args) {
			final String thisClassName = "MainExcelDbTable";
			//
			
			try { //setup the com class with properties file and log file prop key
				acomm = new ACommDbFile(propFileName, args);
				
	  		    for (String thisFileName: acomm.getFileList(Arrays.asList(".xls",".xlst",".txt"))) {
	  		    	   acomm.addPageMsgsLineOut(thisClassName+"====>Process File{" + thisFileName +"}");
	  		    	
						mainApp = new MainAppDataStore(acomm, new ExcelDbTable(), thisFileName, acomm.getFileTextDelimTab());
	  		    	    //mainApp = new MainAppDataStore(acomm, new ExcelDbTable(), thisFileName, '|');
						mainApp.setSourceHeadRowStart(1);
						mainApp.setSourceDataHeadRowStart(3);
						//mainApp.setSourceDataHeadRowEnd(1);
						mainApp.setSourceDataRowStart(4);
						//mainApp.setSourceDataRowEnd(10);
						
						mainApp.doProcess(acomm, "MainExcelDbTable");		
	  		    }
	  		    
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, acomm.getFileTextEndLine(), 3);
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, ' ');
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, acomm.getFileTextDelimTab());
				
				//mainApp.getHtmlServ().outPageLine(acomm, thisClassName+" completed ");
				acomm.end();
				
			} catch (AException e1) {
				throw e1;
			}

		}
//

//
//
// END CLASS
//	
}
