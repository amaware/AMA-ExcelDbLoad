/**
 * 
 */
package net.amaware.apps.exceldbload;

import net.amaware.autil.AFileO;

import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

import org.apache.poi.ss.usermodel.Sheet;

import net.amaware.app.DataStoreReport;
import net.amaware.app.MainAppDataStore;
import net.amaware.aproc.SqlMetaData;
import net.amaware.autil.AComm;
import net.amaware.autil.ACommDb;
import net.amaware.autil.ADataCol;
import net.amaware.autil.ADataColResult;
import net.amaware.autil.ADataColSqlMeta;
import net.amaware.autil.ADatabaseAccess;
import net.amaware.autil.AException;
import net.amaware.autil.AExceptionSql;
import net.amaware.autil.AFileExcelPOI;
import net.amaware.autil.AProperties;
//import net.amaware.serv.DataStore;
import net.amaware.serv.HtmlTargetServ;
import net.amaware.serv.SourceProperty;
//import net.amaware.autil.ASqlStatements;
//import net.amaware.serv.SourceServProperty;

/**
 * @author PSDAA88 - Angelo M Adduci - Sep 6, 2005 3:02:12 PM
 *
 */
public class ExcelDbTable extends DataStoreReport {
	final String thisClassName = this.getClass().getName();
	//
	// Field map
	//*SqlApp AutoGen @2017-04-23 09:26:25.0
//    protected ADataColResult fId = mapDataCol("id");
	/*
    protected ADataColResult fVarcharNull = mapDataCol("VARCHAR_NULL");
    protected ADataColResult fVarcharNotnull = mapDataCol("VARCHAR_NOTNULL");
    protected ADataColResult fIntegerNull = mapDataCol("INTEGER_NULL");
    protected ADataColResult fIntegerNotnull = mapDataCol("INTEGER_NOTNULL");
    protected ADataColResult fDecimalNull = mapDataCol("DECIMAL_NULL");
    protected ADataColResult fDecimalNotnull = mapDataCol("DECIMAL_NOTNULL");
    protected ADataColResult fBooleanNull = mapDataCol("BOOLEAN_NULL");
    protected ADataColResult fBooleanNotnull = mapDataCol("BOOLEAN_NOTNULL");
    protected ADataColResult fTimestampNull = mapDataCol("TIMESTAMP_NULL");
    protected ADataColResult fTimestampNotnull = mapDataCol("TIMESTAMP_NOTNULL");
    protected ADataColResult fDateNull = mapDataCol("DATE_NULL");
    protected ADataColResult fDateNotnull = mapDataCol("DATE_NOTNULL");
    protected ADataColResult fTimeNull = mapDataCol("TIME_NULL");
    protected ADataColResult fTimeNotnull = mapDataCol("TIME_NOTNULL");
    //
    protected ADataColResult fMsgOut = mapDataCol("Msg");
*/
    //
    //DATA_TYPES UDATA_TYPES = new DATA_TYPES();
    //

	//
	String optPropertiesFileName="";
	String optTableName="";
	String optFileDate="";
	
	String optFileOPTION="";
	
	String optRowOPTION="";
	//
	enum OptFileOptions {  None, Insert, Mod;}; 
	OptFileOptions optFileOption = OptFileOptions.None;
	OptFileOptions optFileRowOption = OptFileOptions.None;
	//
	AFileExcelPOI aFileExcelPOI = new AFileExcelPOI(); 
	Sheet aSheetDetail;
	Sheet aSheetResult;
	Sheet aSheetMetaData;
	Sheet aSheetLog;
    //    
    //
    int numRowsIn=0;
    int numRowsToProcess=0;
    int numRowsInserted=0;
    int numRowsDup=0;
    //
    String outFileNamePrefix="";
    String transTS="";
	//
    ADatabaseAccess thisADatabaseAccess;
	List<ADataColSqlMeta> thisADataColSqlMeta = new ArrayList<ADataColSqlMeta>();
	//
	List<ADataColResult> dataHeaderDataColResultList  = new ArrayList<ADataColResult>();;
	//
	//int bypassDataHeadRecCnt=0;
	//
	boolean isDefOut = false;
	//
	public String dbURL = "";
	public String dbTable = "";
	public String dbOptions = "";
	//
	AFileO outSqlInsertFile = new AFileO();
	String outSqlInsertFileNameFull = "";
	//
	int outStatementInsertCtr = 0;
	int outStatementValueCtr = 0;    
    //
	MainAppDataStore thisMainAppDataStore=null;
	//
	/**
	* 
	*/

	public ExcelDbTable() {
		super();

	}

	public void setMainAppDataStore(MainAppDataStore _MainAppDataStore) {
		
		thisMainAppDataStore=_MainAppDataStore;

	}	
	
	@Override
	public DataStoreReport processThis(ACommDb acomm, SourceProperty _aProperty, HtmlTargetServ _aHtmlServ) {
		super.processThis(acomm, _aProperty, _aHtmlServ); // always call this
															// first

		getThisHtmlServ().outPageLine(acomm, thisClassName + "=>processThis");

		_aProperty.displayProperties(acomm);
		
		outFileNamePrefix = acomm.getOutFileDirectoryWithClassName()+AComm.getArgFileName();
		
		//_aProperty.displayProperties(acomm);
		//_aProperty.setValue(SourceProperty.getPropDataRowEnd(), 15);
		//
		transTS=acomm.getCurrTimestampNew();
		//
		acomm.addPageMsgsLineOut(thisClassName
		  + ":processThis StatementId=" + getAStatementsID()
          + " |SourcePropertyFileName=" + _aProperty.getNameFull(acomm)
		 );

		acomm.addPageMsgsLineOut(
				    "        "
				  + " |dbMaxRowsToReturn=" + acomm.getDbRowsMaxReturn()
				  + " |PropertyNumberOfMaxDataRows=" + _aProperty.getValue(SourceProperty.getPropDataRowEnd())
				 );		
		
		
		return this;
	}

	/*
	* 
	* 
	*/

	public String doRequest() {
		return "REPORT-BREAK 1" + "REPORT-BREAK-SUM 2" + "; ";
	}

	@Override
	public boolean doSourceHead(ACommDb acomm, Vector dataFields) throws AException {
		super.doSourceHead(acomm, dataFields);
		// Head Data si not available here.....see doDataHead
		
	    StringBuffer outBuffer = new StringBuffer();
	    outBuffer.append(thisClassName+"=>SourceHeaderCols ");
		
		int optNum=0;
		Vector<String> optionVector = getSourceHeadVector();
		for (String option : optionVector) {
			++optNum;
			
	    	 outBuffer.append(" option#"+optNum+"{" + option + "}");			
			
		}
		
		getThisHtmlServ().outPageLine(acomm, acomm.addPageMsgsLineOut(outBuffer.toString()) , "color:navy;border:solid orange .1em;");
		
		if (optNum < 4) {
			throw new AException(acomm, "File Header must have at least 4 options. #Found{"+optNum+"}");
		}
		
		optPropertiesFileName=optionVector.elementAt(0);
		optTableName=optionVector.elementAt(1);
		optFileDate=optionVector.elementAt(3);
		

		optFileOPTION=optionVector.elementAt(2).toLowerCase();		
		optFileOption = OptFileOptions.None;
		switch (optFileOPTION) {
		case "options":
		case "option":
		case "":
		case "none":	

			break;

		case "insert":	
			optFileOption=OptFileOptions.Insert;
			break;

		case "mod":	
			optFileOption=OptFileOptions.Mod;
			break;
			
			
		default:
			
			throw new AException(acomm, this.getClass().getName()+"=>File header Option invalid{"+optFileOption.toString()+"}"
					            +" |Valid Options{"+Arrays.asList(OptFileOptions.values())+"}"
			                    );
			
			//break;
		} 
		
		getThisHtmlServ().outPageLine(acomm, acomm.addPageMsgsLineOut("=>Using File header Option{"+optFileOption.toString()+"}") , "color:navy;border:solid orange .1em;");
		
		return true;
	}

/*	
*/
     
	@Override
	public boolean doDataHead(ACommDb acomm, int rowNum) throws AException {

		  //if no sql statement on report, comment out next line 
       //setUserTitle2(getThisHtmlServ().formatForSqlout(acomm, getThisStatement()));
       setUserTitle2(thisClassName);
	   //
       /**/
       
		//super.doDataHead(acomm, rowNum);	       
       
       //List<ADataColResult> al = getRowDataColResultList();
       
		Enumeration en = getDataRow().getDataColVec().elements();
		//int cnt=0;
		while (en.hasMoreElements()) {
			ADataColResult aDataColResult = (ADataColResult) en.nextElement();
		    if (aDataColResult != null && 
		    		(aDataColResult.getColumnName().contentEquals("PK-UNIQUE"))) {
    	    	//ADataColResult newADataColResult = new ADataColResult("",aDataColResult.getColumnValue(), aDataColResult.getColumnValue(), true); 
    	    	// dataHeaderDataColResultList.add(newADataColResult);
    	    	//++bypassDataHeadRecCnt;
    	    	if (thisMainAppDataStore!=null) {
    	    	   thisMainAppDataStore.setSourceDataHeadRowStart(3);
    	    	   thisMainAppDataStore.setSourceDataRowStart(4);
		        }
    	    	return true;
	        //} else {
	        //	break;
	        }
		}
		/**/
		super.doDataHead(acomm, rowNum);	
	   //
       setUpDataHead(acomm, rowNum, getRowDataColResultList());
       //			
	   return true;
       //
	}
	
	public void setUpDataHead(ACommDb acomm, int rowNum, List<ADataColResult> _ADataColResultList) throws AException {
		
		int colNum = 0;
		

   		//

		
	     StringBuffer outBuffer = new StringBuffer();
	     outBuffer.append(thisClassName+"=>ColHeaders ");
	     
	     List<String> colHeadList = new ArrayList<String>();
	     //colHeadList.add("Request-Result");
	     for (ADataColResult adcr: _ADataColResultList) {
	    	 
	    	 outBuffer.append(" Name{" + adcr.getColumnName() + "}"
						               + " Title{" + adcr.getColumnTitle() + "}"
						               + " Val{" + adcr.getColumnValue() + "}"
						               );
		    
	    	 adcr.setTableName(dbTable);
	    	 
	    	 colHeadList.add(adcr.getColumnName());
	    	 
		 }		
		
			getThisHtmlServ().outPageLine(acomm, "Data Head Row#" + getSourceRowNum() + " #cols=" + colNum + " line{"
					+ outBuffer.toString() + "}", "color:navy;border:solid orange .1em;");
	     

			//try {
				//C:\projects\amawareData\MainExcelDbTable\output\ExcelDbTable~C:\projects\amawareData\MainExcelDbTable\data_types-2017-04-29.exp-localhost-local.xls}
				
			    String outExcelFileName=AComm.getArgFileFullNameWithClassName().toLowerCase().replace(".xlsx", ".xls");
			    outExcelFileName=outExcelFileName.replace(".xls", ".report.xls");
			    
			    acomm.addPageMsgsLineOut(thisClassName+ "=>Output Excel File Name{" +outExcelFileName +"}");
				
				aFileExcelPOI = new AFileExcelPOI(acomm, outExcelFileName);
				
			//} catch (IOException e) {
			//	throw new AException(acomm, e, "exportFileExcel");
			//}	
			//true to keep connection open till connectionEnd() is done 
			//Need to connectionCommit() for updates...before connectionEnd; 	
			thisADatabaseAccess = new ADatabaseAccess(acomm, optPropertiesFileName, true); //
			//
			thisADataColSqlMeta = thisADatabaseAccess.doDbMetadata(optTableName);				
			//
			//outSqlInsertFileNameFull = acomm.getOutFileDirectoryWithSep()+dbTable+".INSERT.SQL";
			outSqlInsertFileNameFull = acomm.getOutFileDirectoryWithClassName()+optTableName+".INSERT.SQL";
			
			try {
				outSqlInsertFile.openFile(outSqlInsertFileNameFull);
			} catch (IOException e1) {

				throw new AException(acomm, e1, outSqlInsertFileNameFull + " Opened");
			}
			getThisHtmlServ().outPageLine(acomm, "outSqlInsertFile Open{" + outSqlInsertFileNameFull + "}",
					"color:navy;border:solid green 1em;");

			outSqlInsertFile.writeLine("--");
			outSqlInsertFile.writeLine("-- generated by " + thisClassName + " @" + acomm.getCurrTimestampNew());
			outSqlInsertFile.writeLine("-- " + outSqlInsertFileNameFull);
			outSqlInsertFile.writeLine("--");   		
			//
			aSheetDetail = aFileExcelPOI.doCreateNewSheet("Request", 2
 		            , Arrays.asList(optPropertiesFileName,optTableName,optFileOPTION,optFileDate
 		            		,"Ran@"+acomm.getCurrTimestampNew()
 		            		, thisADatabaseAccess.getThisAcomm().getDbUrlDbAndSchemaName()
 		            		)
 		            		
 					, colHeadList
                   );			
			//	
			//
			aSheetMetaData = thisADatabaseAccess.doDbMetadataExcelSheet(aFileExcelPOI,"MetaData");
			//
	   		aSheetLog=aFileExcelPOI.doCreateNewSheet("Log", 2
					  , Arrays.asList("Log")
					  , Arrays.asList("SourceRow#"
					  	        , "Item"
					  	        , "Msg"
					            )
	               );   							
	}
	/*
	* 
	* 
	*/

	@Override
	public boolean doDataRowsNotFound(ACommDb acomm) throws AException {
		super.doDataRowsNotFound(acomm);

		// getThisHtmlServ().outPageLine(acomm, "DataRowsNotFound");

		return true;

	}

	/*
	* 
	* 
	*/

	@Override
	public boolean doDataRow(ACommDb acomm, AException _exceptionSql, boolean _isRowBreak) throws AException {

		// super.doDataRow(acomm, _exceptionSql, _isRowBreak); // sends pout row
		// getThisHtmlServ().outPageLine(acomm, "DataRowFound");

		//int _currRowNum = getSourceRowNum();

		super.doDataRow(acomm, _exceptionSql, _isRowBreak);
		
		//List<ADataColResult> al = getRowDataColResultList();
		
		String sqlInsertStatement="";
	    StringBuffer outBuffer = new StringBuffer();
	    outBuffer.append("=>FileDataRow#{" + getSourceRowNum() + "}");
	    int colNum=0; 
	    String colOptionValue="";
	    for (ADataColResult adcr: getRowDataColResultList()) {
	    	++colNum; 
	    	
	    	if(colNum==1 && adcr.getColumnValue()!=null) {
	    		colOptionValue=adcr.getColumnValue().toLowerCase();	
	    	}
	    	
	    	 outBuffer.append(" |col#{"+colNum+"}"+" Name{" + adcr.getColumnName() + "}"
						               + " Title{" + adcr.getColumnTitle() + "}"
						               + " Val{" + adcr.getColumnValue() + "}"
						               );
		    
	    	 //adcr.setTableName(dbTable);
	    	 
		}		
		
		if (colOptionValue.contentEquals("end")) {
			getThisHtmlServ().outPageLine(acomm, acomm.addPageMsgsLineOut("END requested...No more being processed") , "color:navy;border:solid orange .1em;");
     		aFileExcelPOI.doOutputRowNextBreak(acomm 
 			         , aSheetDetail
 				     , Arrays.asList(
						        this.getClass().getSimpleName() + " at End"
						        //,this.getClass().getSimpleName()
				  	            //, UDATA_TYPES.getInsertStatement(acomm)
					            ) 
			     );  
     		
     		//thisADatabaseAccess.connectionCommit();
     		
			return false;
     		
     		
     		
     		
		} else {
			optRowOPTION=colOptionValue.trim();
			optFileRowOption = OptFileOptions.None;
			switch (optRowOPTION) {
			case "":
			case "none":	

				break;

			case "insert":	
				optFileRowOption=OptFileOptions.Insert;
				break;

			case "mod":	
				optFileRowOption=OptFileOptions.Mod;
				break;
				
				
			default:
				
				throw new AException(acomm, this.getClass().getName()+"=>File rec Option invalid{"+optRowOPTION+"}"
						            +" |Valid Options{"+Arrays.asList(OptFileOptions.values())+"}"
				                    );
				
				//break;
			} 
		}
	
		if (optFileRowOption==OptFileOptions.None) {
			optFileRowOption=optFileOption;	
		}
		
		outBuffer.append("=>Row Option{"+optFileRowOption.name()+"}");
		getThisHtmlServ().outPageLine(acomm, acomm.addPageMsgsLineOut(outBuffer.toString()) , "color:navy;border:solid orange .1em;");		
		
        //
		
		try {
			
			//getThisHtmlServ().outPageRow(acomm, this);
			
			++numRowsToProcess;			
		/*	
		  if ((fVarcharNull.getColumnValue() == null || fVarcharNull.getColumnValue().isEmpty()) 
		  && (fVarcharNotnull.getColumnValue() == null || fVarcharNotnull.getColumnValue().isEmpty())
		     ) 
		  {
			  
				getThisHtmlServ().outPageLineError(acomm,
						acomm.addPageMsgsLineOut(thisClassName + "...BYPASS Row for null fVarcharNull{" + fVarcharNull.getColumnValue() + "}"
								+ " and null fVarcharNotnull{"  + fVarcharNotnull.getColumnValue()  + "}"));			  
			  
			  return true;
		  }
		  */ 

 		  //
			StringBuffer outReqResultBuffer = new StringBuffer();
			List<String> outRowCols = new ArrayList<String>();
			
          
			//if (fAddressZIP.getColumnValue().contains("11228")) {
				   
			//doDSRFieldsValidate(acomm);
			
		    //fTestIntNull.setColumnValue(""); //here because Validate move zero as default to notnum fields
			
			//doDSRFieldsToTableDATA_TYPES(acomm, UDATA_TYPES);

			getThisHtmlServ().outPageRow(acomm, this);

			//
			
			try {
				
				//UDATA_TYPES.doProcessInsertRow(acomm, UDATA_TYPES.getInsertStatement(acomm));
	   	  		
				
				if (optFileRowOption == OptFileOptions.Insert) {
					
					sqlInsertStatement = thisADatabaseAccess.doProcessInsertRow(getRowDataColResultList());
					
					aFileExcelPOI.doOutputRowNext(acomm 
		    			      , aSheetLog, (Arrays.asList(
		    						        ""+getSourceRowNum()
		    						        ,"Row Inserted"
		    				  	            //, UDATA_TYPES.getInsertStatement(acomm)
		    					            ) 
		    			                   )
		    				   );
		   	  		
					getThisHtmlServ().outPageLine(acomm,
							acomm.addPageMsgsLineOut(thisClassName + "...Row Inserted for {" //+ UDATA_TYPES.getInsertStatement(acomm) + "}"
									));

					outReqResultBuffer.append("Inserted");
					
				} else {
					outReqResultBuffer.append("Bypassed");
				}
				
			} catch (AExceptionSql e1) {
				if (e1.isExceptionSqlRowDuplicate(acomm)) { //
					++numRowsDup;
					getThisHtmlServ().outPageLineError(acomm,
							acomm.addPageMsgsLineOut(thisClassName + "...DUP Row NOT Inserted for {" //+ UDATA_TYPES.getInsertStatement(acomm) + "}"
									+ " msg{" + e1.getExceptionMsg() + "}"));
					
		   	  		aFileExcelPOI.doOutputRowNext(acomm 
		    			      , aSheetLog
		    				  , (Arrays.asList(
		    						    ""+getSourceRowNum()
		    						    ,"DUP Row NOT Inserted"
		    				  	        //, UDATA_TYPES.getInsertStatement(acomm)
		    					        ) 
		    			        )
		      		);

		   	  	    outReqResultBuffer.append(e1.getExceptionMsg());
		   	  		
				} else {
					getThisHtmlServ().outPageLine(acomm, 1, "...Row NOT Inserted for {"
							//+ UDATA_TYPES.getInsertStatement(acomm) + "}" + "msg{" + e1.getExceptionMsg() + "}"
							);
					
		   	  		aFileExcelPOI.doOutputRowNext(acomm 
		    			      , aSheetLog
		    				  , (Arrays.asList(
		    						    ""+getSourceRowNum()
		    						    ,"ERROR MSG{" + e1.getExceptionMsg() + "}"
		    				  	        //, UDATA_TYPES.getInsertStatement(acomm)
		    					        ) 
		    			        )
		      		);

		   	  	    outReqResultBuffer.append(e1.getExceptionMsg());
		   	  	
					//throw e1;
				}
			}
			
			//outRowCols.add(outReqResultBuffer.toString());
			
			//String firstField = getDataRowColsToList().get(0);
			//getDataRowColsToList().set(0, "res-"+outReqResultBuffer.toString());
			
			outRowCols.addAll(getDataRowColsToList());
			
			outRowCols.set(0,optFileRowOption.toString()+ "{"+outReqResultBuffer.toString()+ "}");
			
     		aFileExcelPOI.doOutputRowNext(acomm 
  			         , aSheetDetail
  				     , outRowCols
			     );         			
			
			//numRowsInserted += UDATA_TYPES.getPsNumRowsInserted();
			//if (UDATA_TYPES.getPsNumRowsInserted() > 0) {
			//	fMsgOut.setColumnValue("Row Inserted ");
			//	getThisHtmlServ().outPageLine(acomm, 1, "...Row Inserted " // for
			//	);

				//++displayNum;
				//if (displayNum >= displayAtNum) {
				//	acomm.addPageMsgsLineOut(
				//			thisClassName + "=>Row#{" + numRowsIn + "}" + " |#RowsInserted{" + numRowsInserted + "}");
				//	displayNum = 0;
				//}

			//}
			/*	   
			if (numRowsInserted ==3) {
				acomm.addPageMsgsLineOut(thisClassName+"=>Commit @ insert #{" + numRowsInserted +  "}");
				acomm.dbConCommit();
				//throw new AException(acomm, acomm.addPageMsgsLineOut(thisClassName+"=>Abending to see ROLLBACK work @ insert #{" + numRowsInserted +  "}"));
			}
			if (numRowsInserted > 5) {
				acomm.addPageMsgsLineOut(thisClassName+"=>Abending to see ROLLBACK work @ insert #{" + numRowsInserted +  "}");
				throw new AException(acomm, acomm.addPageMsgsLineOut(thisClassName+"=>Abending to see ROLLBACK work @ insert #{" + numRowsInserted +  "}"));
			}
			*/
			ADataColResult aDataColResult;
			int colnum = 0;
			int thisFieldLen = 0;
			int outLineLen = 0;
			int outLineLenMax = 71; // 79;
			int outColInsNum = 0;
			int outColValNum = 0;
			//StringBuffer outLineBuff = new StringBuffer();

			boolean isFirstColOut = false;

			String colResultColNameUse = "";
			String[] stringSplitArray;
			colnum = 0;
			outLineLen = 0;
			//outLineBuff.setLength(0);
			isFirstColOut = false;
			outColInsNum = 0;

			if (sqlInsertStatement.length() > 0) {
				outSqlInsertFile.writeLine("-----------");
				++outStatementInsertCtr;
				outSqlInsertFile.writeLine(" " + sqlInsertStatement);
			}
			//
			
			if (numRowsIn == 100000) {
				acomm.addPageMsgsLineOut(thisClassName+"=>Stop Processng file #{" + numRowsIn +  "}");
				return false;
			}
			
/*			
			if (_currRowNum == 10) {
				getThisHtmlServ().outPageLineCol(acomm,
				    4, "This is Row#" +_currRowNum + " on 4th col");
				getThisHtmlServ().outPageLine(acomm,
					    1, "This is Row#" +_currRowNum + " Line at 1st col");
			   
			}

			if (_currRowNum == 13) {
				getThisHtmlServ().outPageLineCol(acomm,
					    7, "This is Row#" +_currRowNum + " on 7th col");
				   
				}
			
*/			
	    } catch (AExceptionSql e1) {
			throw new AException(acomm, e1, thisClassName
					+ "=>SQL Exception @Row#" + numRowsIn
			        + " |SQLCode{" + e1.getExceptionCode() + "}"
			        + " |SQLMsg{" + e1.getExceptionMsg() + "}"
			        );
			
		} catch (Exception e) {
			throw new AException(acomm, e, thisClassName
					+ "=>doSelectRowCurr=>Row#" + numRowsIn);
			
		}
		
		
		
		
		if (!isDataRowOut()) {
			  super.doDataRow(acomm, _exceptionSql, _isRowBreak);
		}
		
		return true; // or false to stop processing of file

	}

	/**
	 * @return
	 */
	@Override
	public boolean doDataRowBreak(ACommDb acomm) throws AException {

		int _currRowNum = getDataRowNum();
		/*
		 * aHtmlServ.outPageTableLine(acomm, 1, "RowBreak Pre at Row#"
		 * +_currRowNum);
		 * 
		 * aHtmlServ.outRowBreak(acomm, this);
		 * 
		 * aHtmlServ.outPageTableLine(acomm, 1, "RowBreak Post at Row#"
		 * +_currRowNum);
		 */

		return super.doDataRowBreak(acomm);
	}

	@Override
	public boolean doDataRowsEnded(ACommDb acomm) throws AException {

		

		int lineNum=0,colNum=0,colNumMax=10;
		StringBuffer outLineBuff = new StringBuffer();
		//
	    //
		getThisHtmlServ().outPageLine(acomm, "Source File Ended",htmlTitleStyle);
		//
		//
		acomm.addPageMsgsLineOut("________________________________" 
				);
		acomm.addPageMsgsLineOut(thisClassName
				    +":doDataRowsEnded=>StatementId=" + getAStatementsID() 
               		+ " @SourceRows#" + getSourceRowNum()
               		+ " @DataRow#" + getDataRowNum()
    				+ " |#MaxRows=" + getSourceDataRowEndNum()
    				);
		
        if (getSourceRowNum() >  getSourceDataRowEndNum()) {
 	    	getThisHtmlServ().outPageLineWarning(acomm,  
	    			 "More data from Source may exist....ended due to requested max rows"
               		+ " |SourceRows#" + getSourceRowNum()
               		+ " |DataRow#" + getDataRowNum()
    				+" |#MaxRows=" + getSourceDataRowEndNum()
 	    			
					);
 	    	
 	    	super.doDataRowsEnded(acomm);
 	    	throw new AException(acomm, "Requested MAX ROWS EXCEEDED...MORE ROWS EXIST" 
               		+ " @SourceRows#" + getSourceRowNum()
               		+ " @DataRow#" + getDataRowNum()
    				+" |#MaxRows=" + getSourceDataRowEndNum()
 	    					);
         }
        //
		outSqlInsertFile.writeLine("--____________________________________________" );
        
		outSqlInsertFile
		//.writeLine("-- #Insert cols out{" + outColInsNum + "}" + " #value cols out{" + outColValNum + "}");
		.writeLine("-- #Insert rows{" + outStatementInsertCtr + "}" );
		outSqlInsertFile.writeLine("--____________________________________________" );

		outSqlInsertFile.closeFile();
        //
		aFileExcelPOI.doOutputRowNext(acomm 
			      , aSheetLog
				  , (Arrays.asList(""+getSourceRowNum()
						    , "At End"
						    , "#SummaryRows{"+aSheetDetail.getLastRowNum() +"}"
						    //  + " |#DetailRows{"+aSheetDetail.getLastRowNum() +"}"
					        ) 
			        )
		);   
  		
	    thisADatabaseAccess.doQueryRsExcel(aFileExcelPOI
	            , "Results"
	            , "Select *"
	        +" from " + optTableName  
	     //+ " Where field_nme  = '" + ufieldname +"'" 
	        
	     //+ " order by entry_type, entry_subject, entry_topic"
	     );  		
  		//
	    thisADatabaseAccess.connectionCommit();
	    thisADatabaseAccess.connectionEnd();
	    //
   		try {
			aFileExcelPOI.doOutputEnd(acomm);
		} catch (IOException e) {
			throw new AException(acomm, e, " Close of outFileExcel");
		}
        
		return super.doDataRowsEnded(acomm); // or false to stop processing of file

	}
	//
	

	//
	//* SqlApp DataStoreReport SET Table columns from DSR
	//
	//*SqlApp AutoGen @2017-04-23 09:26:25.0
	 public void doDSRFieldsToTableDATA_TYPES(ACommDb acomm) {
	            // doDSRFieldsToTableDATA_TYPES(acomm, UDATA_TYPES);
	//
	 } //End doDSRFieldsToTableDATA_TYPES qDATA_TYPES
	//
	//
	 /*
	 public void doDSRFieldsToTableDATA_TYPES(ACommDb acomm, DATA_TYPES _qClass) {
	   //_qClass.setId(fId.getColumnValue());
	   _qClass.setVarcharNull(fVarcharNull.getColumnValue());
	   _qClass.setVarcharNotnull(fVarcharNotnull.getColumnValue());
	   _qClass.setIntegerNull(fIntegerNull.getColumnValue());
	   _qClass.setIntegerNotnull(fIntegerNotnull.getColumnValue());
	   _qClass.setDecimalNull(fDecimalNull.getColumnValue());
	   _qClass.setDecimalNotnull(fDecimalNotnull.getColumnValue());
	   _qClass.setBooleanNull(fBooleanNull.getColumnValue());
	   _qClass.setBooleanNotnull(fBooleanNotnull.getColumnValue());
	   _qClass.setTimestampNull(fTimestampNull.getColumnValue());
	   _qClass.setTimestampNotnull(fTimestampNotnull.getColumnValue());
	   _qClass.setDateNull(fDateNull.getColumnValue());
	   _qClass.setDateNotnull(fDateNotnull.getColumnValue());
	   _qClass.setTimeNull(fTimeNull.getColumnValue());
	   _qClass.setTimeNotnull(fTimeNotnull.getColumnValue());
	//
	 } //End doDSRFieldsToTable DATA_TYPES _qClass
	*/
	
	//
	//
	//* SqlApp DataStoreReport SET Data Fields from Table columns
	//
	//*SqlApp AutoGen @2017-04-23 09:26:25.0
	 public void doDSRFieldsFromTableDATA_TYPES(ACommDb acomm) {
	           //  doDSRFieldsFromTableDATA_TYPES(acomm,  UDATA_TYPES);
	//
	 } //End doDSRFieldsFromTableDATA_TYPES
	//
	 /*
	 public void doDSRFieldsFromTableDATA_TYPES(ACommDb acomm, DATA_TYPES _qClass) {
	    //fId.setColumnValue(_qClass.getId());
	    fVarcharNull.setColumnValue(_qClass.getVarcharNull());
	    fVarcharNotnull.setColumnValue(_qClass.getVarcharNotnull());
	    fIntegerNull.setColumnValue(_qClass.getIntegerNull());
	    fIntegerNotnull.setColumnValue(_qClass.getIntegerNotnull());
	    fDecimalNull.setColumnValue(_qClass.getDecimalNull());
	    fDecimalNotnull.setColumnValue(_qClass.getDecimalNotnull());
	    fBooleanNull.setColumnValue(_qClass.getBooleanNull());
	    fBooleanNotnull.setColumnValue(_qClass.getBooleanNotnull());
	    fTimestampNull.setColumnValue(_qClass.getTimestampNull());
	    fTimestampNotnull.setColumnValue(_qClass.getTimestampNotnull());
	    fDateNull.setColumnValue(_qClass.getDateNull());
	    fDateNotnull.setColumnValue(_qClass.getDateNotnull());
	    fTimeNull.setColumnValue(_qClass.getTimeNull());
	    fTimeNotnull.setColumnValue(_qClass.getTimeNotnull());
	//
	 } //End doDSRFieldsFromTable qDATA_TYPES
	*/
	//
	//
	//* SqlApp DataStoreReport Validate input  Data Fields
	//
	//*SqlApp AutoGen @2017-04-23 09:26:25.0
	 public void doDSRFieldsValidate(ACommDb acomm) throws Exception {
	   // fId.setColumnValue(String.valueOf(doFieldValidateInt(acomm, fId.getColumnValue(),0)));
		 /*
	    fVarcharNull.setColumnValue(doFieldValidateString(acomm, fVarcharNull.getColumnValue()));
	    fVarcharNotnull.setColumnValue(doFieldValidateString(acomm, fVarcharNotnull.getColumnValue()));
	    fIntegerNull.setColumnValue(String.valueOf(doFieldValidateInt(acomm, fIntegerNull.getColumnValue(),0)));
	    fIntegerNotnull.setColumnValue(String.valueOf(doFieldValidateInt(acomm, fIntegerNotnull.getColumnValue(),0)));
	    fDecimalNull.setColumnValue(String.valueOf(doFieldValidateNum(acomm, fDecimalNull.getColumnValue(),0.0)));
	    fDecimalNotnull.setColumnValue(String.valueOf(doFieldValidateNum(acomm, fDecimalNotnull.getColumnValue(),0.0)));
	    fBooleanNull.setColumnValue(doFieldValidateBoolean(acomm, fBooleanNull.getColumnValue()));
	    fBooleanNotnull.setColumnValue(doFieldValidateBoolean(acomm, fBooleanNotnull.getColumnValue()));
	    fTimestampNull.setColumnValue(doFieldValidateString(acomm, fTimestampNull.getColumnValue()));
	    fTimestampNotnull.setColumnValue(doFieldValidateString(acomm, fTimestampNotnull.getColumnValue()));
	    
	    try {
	    fDateNull.setColumnValue(doFieldValidateDate(acomm, fDateNull.getColumnValue()));
	    } catch (Exception e) { throw e; }
	    
	    try {
	    fDateNotnull.setColumnValue(doFieldValidateDate(acomm, fDateNotnull.getColumnValue()));
	    } catch (Exception e) { throw e; }
	    
	    fTimeNull.setColumnValue(doFieldValidateString(acomm, fTimeNull.getColumnValue()));
	    fTimeNotnull.setColumnValue(doFieldValidateString(acomm, fTimeNotnull.getColumnValue()));
	//
	  */
	 } //End doDSRFieldsValidate
	 
	//
	//
	// END
	//
}
