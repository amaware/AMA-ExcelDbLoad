package net.amaware.apps.exceldbload;
import java.io.FileNotFoundException;
import java.io.IOException;
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
	  		    	   acomm.addPageMsgsLineOut(thisClassName+"=>Input File{" + thisFileName +"}");
	  		    	   //
						int colHeadRow=2;
						int colDetRow=3;
						//
						try {
						 AFileExcelPOI aFileExcelPOI = new AFileExcelPOI();
						
						 aFileExcelPOI.openInput(thisFileName);
						 
						 List<String> excelColList = aFileExcelPOI.readNextAsStringList();
						 int rowNum=0;
						 while (excelColList.size() > 0 && rowNum<3) {
							 ++rowNum;
							 if (excelColList.get(0).contentEquals("PK-UNIQUE")) {
								colHeadRow=3;
								colDetRow=4;
								break; 
							 }
							 excelColList = aFileExcelPOI.readNextAsStringList();
						 }
						 
			            } catch (IOException e) {
			            	throw new AException(acomm, thisClassName + " File Not found{"+thisFileName+"}");
		                }	  		    	   
	  		    	   
	  		    	   //
						
						acomm.addPageMsgsLineOut(thisClassName+"==>File ColHeaderRow{" + colHeadRow +"}"
								                +" |ColDetailRowStart{" + colDetRow +"}");
						
	  		    	    ExcelDbTable aExcelDbTable = new ExcelDbTable();
	  		    	    
						mainApp = new MainAppDataStore(acomm, aExcelDbTable, thisFileName, acomm.getFileTextDelimTab());
	  		    	    //mainApp = new MainAppDataStore(acomm, new ExcelDbTable(), thisFileName, '|');
						//
						mainApp.setSourceHeadRowStart(1);

						//
						mainApp.setSourceDataHeadRowStart(colHeadRow);
						mainApp.setSourceDataRowStart(colDetRow);
						//
						//mainApp.setSourceDataRowEnd(10);
						
						aExcelDbTable.setMainAppDataStore(mainApp);
						
						//
						mainApp.doProcess(acomm, "MainExcelDbTable");		
	  		    }
	  		    
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, acomm.getFileTextEndLine(), 3);
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, ' ');
				//mainApp = new MainAppDataStore(acomm, _fileProcess, args, acomm.getFileTextDelimTab());
				
				//mainApp.getHtmlServ().outPageLine(acomm, thisClassName+" completed ");
				acomm.end();
				
			} catch (AException e1) {
				acomm.addPageMsgsLineOut("MainExcelDbLoad AException msg{"+e1.getMessage()+"}");
				throw e1;
			}

		}
//
//

//
// END CLASS
//	
}
