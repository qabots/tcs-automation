
var overAllTestCaseStatus,listOfEnvironments = Sys.OleObject("Scripting.Dictionary");
var gLogPath,strTestCaseHTMLFilePath,S_NO,htmlreportcontent='';
var gTestCasePicPath,gPictureLogPath,executedEnvs;
var PASS_COLR = "Green", FAIL_COLOR = "Red";
var totalTCCount = 0,passTCCount = 0,failTCCount = 0,exeStartTime,exeEndTime,tcStartTime,tcEndTime,strHTMLHighLevelReport;

//set ProjectSuite Varaibles for Report
function setProjectSuiteLevelVariables(){
  if (!ProjectSuite.Variables.VariableExists("exeStartTime"))
     ProjectSuite.Variables.AddVariable("exeStartTime","String");
  if (!ProjectSuite.Variables.VariableExists("gPictureLogPath"))
     ProjectSuite.Variables.AddVariable("gPictureLogPath","String");
  if (!ProjectSuite.Variables.VariableExists("strTestCaseHTMLFilePath"))
     ProjectSuite.Variables.AddVariable("strTestCaseHTMLFilePath","String");
  if (!ProjectSuite.Variables.VariableExists("totalTCCount"))
     ProjectSuite.Variables.AddVariable("totalTCCount","Integer");
  if (!ProjectSuite.Variables.VariableExists("passTCCount"))
     ProjectSuite.Variables.AddVariable("passTCCount","Integer");
  if (!ProjectSuite.Variables.VariableExists("failTCCount"))
     ProjectSuite.Variables.AddVariable("failTCCount","Integer");
}
//set the ReportLogPath
function setLogsPath(str_PathToLogs)
{
      if(aqString.SubString(str_PathToLogs,(aqString.GetLength(str_PathToLogs)-1),aqString.GetLength(str_PathToLogs)) != "\\")
      {
            gLogPath =  str_PathToLogs + "\\";         
      }
      else
      {
            gLogPath =  str_PathToLogs;
      }
          
}
//Sets StartTime of test 
function setExecutionStartTime(time_ExecutionStart)
{
      ProjectSuite.Variables.exeStartTime = time_ExecutionStart;
}
//Sets EndTime of test 
function setExecutionEndTime(time_ExecutionEnd)
{
      exeEndTime = time_ExecutionEnd;   
}
//create the report folder
function fn_createreportfolder()
{
  //Log.Message("Project Suite Path : " + ProjectSuite.Path);

  gLogPath = ProjectSuite.Path + "Results\\" + fn_getprojectSuiteName() + "_" + fn_getuniquefilename( )+ "\\";
  //ProjectSuite.Variables.gLogPath = gLogPath;
  //Log.Message(gLogPath);
}
//TestCase HTML Reports
//#########################################################################################################################################
//Create a report file 

function fn_createreportfile(executedEnvsStr){
  executedEnvs = executedEnvsStr;
  setProjectSuiteLevelVariables();
  fn_createtestcasereportfile();
  fn_createreportheader(executedEnvs)
  fn_createtestcasetable();
}

//Create report folder and filename
function fn_createtestcasereportfile()
{
      //set execution start time 
      setExecutionStartTime(aqDateTime.Now());
      if (strTestCaseHTMLFilePath== undefined){
        var tempunquePath = fn_getuniquefilename();
        fn_createreportfolder();
        Log.Message(gLogPath);

        aqFileSystem.CreateFolder(gLogPath);
        
        //        
        gPictureLogPath = gLogPath + "Picture";
      aqFileSystem.CreateFolder(gPictureLogPath);
      gPictureLogPath = gPictureLogPath + "\\";
      ProjectSuite.Variables.gPictureLogPath = gPictureLogPath;
      
      gTestCasePicPath = "\\Picture\\";
      
	    strTestCaseHTMLFilePath = gLogPath + fn_getprojectSuiteName() + "-" + tempunquePath + ".htm";
      ProjectSuite.Variables.strTestCaseHTMLFilePath = strTestCaseHTMLFilePath;

      }

}
//Create report header
function fn_createreportheader (executedEnvs){
  var htmlreport; 

  if (exeEndTime == undefined ){ setExecutionEndTime(aqDateTime.Now());}
  
  htmlreport = fn_starthtmlbodycolor("White") + fn_htmlnewline() + fn_fontstart() + fn_header("Automation Highlevel Report"); 
	htmlreport = htmlreport + fn_htmlstarttable(1) + fn_htmlcreateheaders("Project Suite Name|Executed Environments|Machine Name") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_getprojectSuiteName(),"Center","White") ;
	//htmlreport = htmlreport + fn_fontcolor("White") + "<b>" + fn_htmlrowdata( "tempOverAllStatus","Center","White")  + "</b>" + fn_fontcolor("Black") ;
	htmlreport = htmlreport + fn_htmlrowdata(executedEnvs,"Center","#F5F5DC") ;
	htmlreport = htmlreport + fn_htmlrowdata(Sys.HostName,"Center","White") +   fn_htmlendrow() + fn_htmlclosetable();
	
	htmlreport = htmlreport +fn_htmlnewline() + fn_htmlnewline() +  fn_htmlstarttable(1) + fn_htmlcreateheaders("Execution Date|Start Time|End Time|Total Duration") ;
  
	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdate(ProjectSuite.Variables.exeStartTime),"Center","White") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdatetime(ProjectSuite.Variables.exeStartTime) ,"Center","White") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdatetime(exeEndTime),"Center","White") ;  	
  
	htmlreport = htmlreport + fn_htmlrowdata(fn_gettimediff(exeEndTime,ProjectSuite.Variables.exeStartTime),"Center","White") ;
	htmlreport = htmlreport + fn_htmlendrow() + fn_htmlclosetable()  + fn_htmlnewline() +  fn_htmlnewline() ;
	
	htmlreport = htmlreport  + "<center>" + fn_initializeCricle("total",ProjectSuite.Variables.totalTCCount) + fn_initializeCricle("Pass",ProjectSuite.Variables.passTCCount) + fn_initializeCricle("Fail",ProjectSuite.Variables.failTCCount)
	htmlreport = htmlreport  + fn_htmlnewline() +    fn_htmlnewline() + fn_htmlnewline()  + fn_htmlnewline() + fn_htmlnewline() + "</center>";
  htmlreport = aqString.Replace(htmlreport,"undefined","-");
  htmlreportcontent = htmlreport;
  
   
      
}

//Create testcase table header
function fn_createtestcasetable (){
  var htmlreport = "                                                                                                                          "; 
  htmlreport =  htmlreport + fn_htmlnewline()+fn_htmlnewline()+fn_fontstart() + fn_htmlstarttable(1) + fn_htmlcreateheaders("Step No.|Actions|Expected Result|Actual Result|Status|Picture ") ;
  htmlreportcontent = htmlreportcontent + htmlreport;

}
//Create projectname row
function fn_createprojectnamerow (prjname){
  var prjrow; 
  prjrow = fn_htmlstartrow() + fn_htmlmergerow("<B>Project Name:</B> " + prjname,"Left","White")+ fn_htmlendrow();
  htmlreportcontent = htmlreportcontent + prjrow;
  S_NO=1;
}
//Returns  Test Step
function fn_createteststep(strStatus,strRequiredAction, strExpected,strActual,wantScreenShot)
{
	var tempCode,statusColor,strImgLink="";
	
  switch (strStatus) 
	{
		case 1://PASS
			statusColor = PASS_COLR;
			strStatus = "Pass";
			Log.Checkpoint(strActual,strExpected);
      ProjectSuite.Variables.passTCCount = ProjectSuite.Variables.passTCCount + 1;
			break;
		case 0: //FAIL
			statusColor = FAIL_COLOR;
			strStatus = "Fail";
			strImgLink = fn_getscreenshotlink();
			Log.Error(strActual,strExpected);
      ProjectSuite.Variables.failTCCount = ProjectSuite.Variables.failTCCount + 1
			break;			
	} 
   ProjectSuite.Variables.totalTCCount = ProjectSuite.Variables.totalTCCount + 1;    
	if(wantScreenShot)
      {
            strImgLink = fn_getscreenshotlink();          
      }
      
	tempCode = fn_htmlstartrow() + fn_htmlrowdata(S_NO,"Center","White") ;
  tempCode = tempCode + fn_htmlrowdata(strRequiredAction,"Left","White") ;
	tempCode = tempCode + fn_htmlrowdata(strExpected,"Left","White") ;
	tempCode = tempCode + fn_htmlrowdata(strActual,"Left","White") ;
	tempCode = tempCode + fn_htmlrowdata(strStatus,"Center",statusColor) ;
	
	if (strImgLink != "")
	{
		tempCode = tempCode + fn_htmlrowdata(fn_createlink(gTestCasePicPath) ,"Center","White") ;
	}
	else
	{
		tempCode = tempCode + fn_htmlrowdata("-" ,"Center","White") ;
	}
      tempCode = tempCode +  fn_htmlendrow();
  //    currentTestCaseStatus = currentTestCaseStatus + "," + strStatus;
	 overAllTestCaseStatus = overAllTestCaseStatus + "," + strStatus;
      
  htmlreportcontent = htmlreportcontent +tempCode;
	S_NO = S_NO + 1;
 fn_completetestcase();
}

//Consolidate the Test Case Overview, run details and Test Steps
function fn_completetestcase()
{
	var strforHTMLReport;
	var createFile=true;
  var reportFile;

  strforHTMLReport = htmlreportcontent;
	strforHTMLReport = aqString.Replace(strforHTMLReport,"undefined","-");
  if (aqFile.Exists(ProjectSuite.Variables.strTestCaseHTMLFilePath)){
   createFile=false;    
   }
    
	if(aqFile.WriteToTextFile(ProjectSuite.Variables.strTestCaseHTMLFilePath,strforHTMLReport,20,createFile))
  {
            Log.Message("Log created for project true");
      }
      else
      {
            Log.Message("Log created for project false");     
      }

  htmlreportcontent = "";
     
}

//#########################################################################################################################################
function fn_generatehighlevelreport(executedEnvs)
{
//	#F5DEB3,#F5F5DC
	var htmlreport,strModuleName = "",strExecutedEnvs = "";
      setExecutionEndTime(aqDateTime.Now());
      if (executedEnvs != undefined){
        strExecutedEnvs = executedEnvs;
      }
      if(aqString.Find(overAllTestCaseStatus,"Fail",1) != -1)
      { 
            tempOverAllStatus = "Fail";
            tempOverAllstatusColor = FAIL_COLOR;
      }
      else if(aqString.Find(overAllTestCaseStatus,"Pass",1) != -1)
      {
            tempOverAllStatus = "Pass";
            tempOverAllstatusColor = PASS_COLR;
      }

      
 /*  
	htmlreport = fn_starthtmlbodycolor("White") + fn_htmlnewline() + fn_fontstart() + fn_header("Automation Highlevel Report"); 
	htmlreport = htmlreport + fn_htmlstarttable(1) + fn_htmlcreateheaders("Module Name(s)|OverAll Status|Executed Environments|Machine Name") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_getprojectSuiteName(),"Center","White") ;
	htmlreport = htmlreport + fn_fontcolor("White") + "<b>" + fn_htmlrowdata( tempOverAllStatus,"Center",tempOverAllstatusColor)  + "</b>" + fn_fontcolor("Black") ;
	htmlreport = htmlreport + fn_htmlrowdata(executedEnvs,"Center","#F5F5DC") ;
	htmlreport = htmlreport + fn_htmlrowdata(Sys.HostName,"Center","White") +   fn_htmlendrow() + fn_htmlclosetable();
	
	htmlreport = htmlreport +fn_htmlnewline() + fn_htmlnewline() +  fn_htmlstarttable(1) + fn_htmlcreateheaders("Execution Date|Start Time|End Time|Total Duration") ;

	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdate(ProjectSuite.Variables.exeStartTime),"Center","White") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdatetime(ProjectSuite.Variables.exeStartTime) ,"Center","White") ;
	htmlreport = htmlreport + fn_htmlrowdata(fn_formatdatetime(exeEndTime),"Center","White") ;  	

	htmlreport = htmlreport + fn_htmlrowdata(fn_gettimediff(exeEndTime,ProjectSuite.Variables.exeStartTime),"Center","White") ;
	htmlreport = htmlreport + fn_htmlendrow() + fn_htmlclosetable()  + fn_htmlnewline() +  fn_htmlnewline() ;
	
	htmlreport = htmlreport  + "<center>" + fn_initializeCricle("total",ProjectSuite.Variables.totalTCCount) + fn_initializeCricle("Pass",ProjectSuite.Variables.passTCCount) + fn_initializeCricle("Fail",ProjectSuite.Variables.failTCCount)
	htmlreport = htmlreport  + fn_htmlnewline() +    fn_htmlnewline() + fn_htmlnewline() + fn_htmlnewline() + fn_htmlnewline() + "</center>";*/
  
  fn_createreportheader (executedEnvs);   
  
   // Log.Message("**************test filee path : " + ProjectSuite.Variables.strTestCaseHTMLFilePath + "********" );
   var myFile = aqFile.OpenTextFile(ProjectSuite.Variables.strTestCaseHTMLFilePath, aqFile.faReadWrite, 22);

    myFile.Cursor = 0;  
    myFile.WriteLine(htmlreportcontent);

    myFile.Close();
    
    htmlreportcontent = "";
    ProjectSuite.Variables.exeStartTime = "";
    ProjectSuite.Variables.failTCCount= 0, ProjectSuite.Variables.passTCCount= 0, ProjectSuite.Variables.totalTCCount = 0;
    ProjectSuite.Variables.gPictureLogPath="";ProjectSuite.Variables.strTestCaseHTMLFilePath="";

}

/*#########################################################################################################################################
*Support Libraries
#########################################################################################################################################*/
//Returns the HTML tag 
function fn_starthtmlbodycolor(strbodyColor)
{
      return "<html> <body bgcolor = " + strbodyColor + ">" + Chr(10);
}
//Returns the HTML end tag
function fn_endhtml()
{
      return " </body> </html>" + + Chr(10);
}
//Returns  HTML Table
function fn_htmlstarttable(intTableType)
{
      var tempCode;
	switch ( intTableType)
 	{
   	      case 1:
		      tempCode = "<table width=" + Chr(34) +  "80%" +  Chr(34) + "cellspacing=" +Chr(34) +  "2" +  Chr(34)  + " cellpadding=" + Chr(34) + "0" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + " align=" + Chr(34) + "center" + Chr(34) + " bgcolor=" + Chr(34) + "#556B2F" + Chr(34) + " font face=" + Chr(34) + "Calibri" + Chr(34) +">" ;
			gcurtblbordercolor = "#556B2F";
			return tempCode + Chr(10);	
     			break;
   		case 2:
 			tempCode = "<table width=" + Chr(34) +  "80%" +  Chr(34) + "cellspacing=" +Chr(34) +  "2" +  Chr(34)  + " cellpadding=" + Chr(34) + "0" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + " align=" + Chr(34) + "center" + Chr(34) + " bgcolor=" + Chr(34) + "#191970" + Chr(34) + " font face=" + Chr(34) + "Calibri" + Chr(34) +">";
			gcurtblbordercolor = "#191970";
			return tempCode + Chr(10);
     			break;
		case 3:
 			tempCode = "<table width=" + Chr(34) +  "80%" +  Chr(34) + "cellspacing=" +Chr(34) +  "2" +  Chr(34)  + " cellpadding=" + Chr(34) + "0" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + " align=" + Chr(34) + "center" + Chr(34) + " bgcolor=" + Chr(34) + "#000000" + Chr(34) + " font face=" + Chr(34) + "Calibri" + Chr(34) +">";
			gcurtblbordercolor = "#000000";
			return tempCode + Chr(10);
     			break;
   		default:
			tempCode = "<table width=" + Chr(34) +  "80%" +  Chr(34) + "cellspacing=" +Chr(34) +  "2" +  Chr(34)  + " cellpadding=" + Chr(34) + "0" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + " align=" + Chr(34) + "center" + Chr(34) + " bgcolor=" + Chr(34) + "#2F4F4F" + Chr(34) + " font face=" + Chr(34) + "Calibri" + Chr(34) +">";
			gcurtblbordercolor = "#2F4F4F";		
			return tempCode + Chr(10);
 	}
		  
}
//Closing the HTML Table
function fn_htmlclosetable()
{
	var tempCode;
	tempCode = "</table>";
	return tempCode + Chr(10);
}
//Returns  HTML headers
function fn_htmlcreateheaders(strHTMLHeaders)
{
	var tempCode;
	aqString.ListSeparator = "|" ;
	if(strHTMLHeaders != "")
	{
	      tempCode = "<thead align = " + chr(34) + "center" + chr(34) + " bgcolor = " + chr(34) + gcurtblbordercolor + chr(34) + "> <tr> ";
		for(var i = 0; i < aqString.GetListLength(strHTMLHeaders) ; i++)
            {
		      tempCode = tempCode + "  <th>" + fn_fontcolor("White") + aqString.GetListItem(strHTMLHeaders,i) + "</th>";
            }
		tempCode = tempCode + " </tr> </thead>";
		return tempCode + Chr(10);
	}
	else
	{
		return "";
	} 
}
//Starting code of theHTML row 
function fn_htmlstartrow()
{ 
      return "<tr>" + Chr(10); 
}

//Merge Row 
function fn_htmlmergerow(strHTMLRowData,stralign,bgcolor)
{ 
  var tempCode;
	if(strHTMLRowData == "" )
	{ 
            strHTMLRowData = "-";
      }
	tempCode = "<td colspan='6' bgcolor = " + chr(34) + "f88379" + chr(34) + "align = " + chr(34) + stralign + chr(34) + ">" + strHTMLRowData + "</td>";
	return tempCode + Chr(10);
}
//Ending code of theHTML row
function fn_htmlendrow()
{ 
      return "</tr>" + Chr(10); 
}
//Inserting row data into HTML table
function fn_htmlrowdata(strHTMLRowData,stralign,bgcolor)
{
	var tempCode;
	if(strHTMLRowData == "" )
	{ 
            strHTMLRowData = "-";
      }
	tempCode = "<td bgcolor = " + chr(34) + bgcolor + chr(34) + "align = " + chr(34) + stralign + chr(34) + ">" + strHTMLRowData + "</td>";
	return tempCode + Chr(10);
}
//Returns  Font tag with default font as Calibri
function fn_fontstart()
{ 
      return "<font face = " + chr(34) +"Calibri" + chr(34) + ">"  + Chr(10);
}
//Ending the Font limit
function fn_fontend()
{ 
      return "</font>" + Chr(10); 
}
//Changing thefont color 
function fn_fontcolor(strColorName)
{
	var tempCode;
	tempCode = "<font color = " + chr(34) + strColorName + chr(34) + ">";
	return tempCode + Chr(10);
}
//Returns  HTML Header data
function fn_header(strValue)
{
	return "<H2> <center> <font color = " + chr(34) + "#2F4F4F" + chr(34) + ">" + strValue +"</center></H2>" + Chr(10);
}
//Returns  HTML new line
function fn_htmlnewline()
{
	return "<br>" + Chr(10);
}
//Returns  HTML hyber link
function fn_createlink(strImgLink)
{
	var tempCode;
 	tempCode = "<a href = " + chr(34) + strImgLink + chr(34) + "> View </a>";
	return tempCode + Chr(10);
}
function fn_initializeCricle(strStatus,intCount)
{
	var tempCode;
  if (intCount === 0 || aqString.GetLength(intCount)== 1 ){
       intCount = " "+intCount;
  }
	switch(strStatus)
	{
		case "Pass":
			tempCode = "   Passed  " + "<style type=" + chr(34) + "text/css" + chr(34) + ">" + ".passCircle {    display:inline-block;    border-radius:50%;    border:2px solid;  border-color:Green; font-size:32px;}"
			tempCode = tempCode + ".passCircle:before,.passCircle:after {    content:'\\200B';    display:inline-block;    line-height:0px;    padding-top:50%;    padding-bottom:50%;}"
			tempCode = tempCode + ".passCircle:before {    padding-left:8px;}.passCircle:after {    padding-right:8px;} </style>"
			tempCode = tempCode + "<span class=" + chr(34) + "passCircle" + chr(34) + ">" + intCount +"</span>"
			return tempCode + Chr(10);
			break;
		case "Fail":
			tempCode ="  Failed  " +  "<style type=" + chr(34) + "text/css" + chr(34) + ">" + ".failCircle {    display:inline-block;    border-radius:50%;    border:2px solid;  border-color: Red; font-size:32px;}"
			tempCode = tempCode + ".failCircle:before,.failCircle:after {    content:'\\200B';    display:inline-block;    line-height:0px;    padding-top:50%;    padding-bottom:50%;}"
			tempCode = tempCode + ".failCircle:before {    padding-left:8px;}.failCircle:after {    padding-right:8px;} </style>"			
			tempCode = tempCode + "<span class=" + chr(34) + "failCircle" + chr(34) + ">" + intCount +"</span>"
			return tempCode + Chr(10);
			break;	
    case "total":
			tempCode = "  Total Test Case(s)  " + "<style type=" + chr(34) + "text/css" + chr(34) + ">" + ".totalCircle {    display:inline-block;    border-radius:50%;    border:2px solid;  border-color: Blue; font-size:36px;}"
			tempCode = tempCode + ".totalCircle:before,.totalCircle:after {    content:'\\200B';    display:inline-block;    line-height:0px;    padding-top:50%;    padding-bottom:50%;}"
			tempCode = tempCode + ".totalCircle:before {    padding-left:8px;}.totalCircle:after {    padding-right:8px;} </style>"			
			tempCode = tempCode + "<span class=" + chr(34) + "totalCircle" + chr(34) + ">" + intCount +"</span>"
			return tempCode + Chr(10);
			break;				
		}	
}

// Caputring Screenshot
function fn_getscreenshotlink()
{
      gTestCasePicPath = "Picture\\";
	var tempObject;
      var tempuniquename = fn_getuniquefilename();
	tempObject = ProjectSuite.Variables.gPictureLogPath + "Fail_" + tempuniquename + ".png";
      gTestCasePicPath = gTestCasePicPath + "Fail_" + tempuniquename + ".png";
	Sys.Desktop.ActiveWindow().Picture().SaveToFile(tempObject);
	return tempObject;	
}
//Returns the time difference between two time : Format 12 Hrs 30 Mins 45 Sec
function fn_gettimediff(endDate,startDate)
{
try
{
	var hrDiff,miDiff,secDiff;
	hrDiff = aqDateTime.GetHours(aqDateTime.TimeInterval(aqConvert.StrToDateTime(startDate),aqConvert.StrToDateTime(endDate)));
	miDiff = aqDateTime.GetMinutes(aqDateTime.TimeInterval(aqConvert.StrToDateTime(startDate),aqConvert.StrToDateTime(endDate)));
	secDiff = aqDateTime.GetSeconds(aqDateTime.TimeInterval(aqConvert.StrToDateTime(startDate),aqConvert.StrToDateTime(endDate)));

	return hrDiff + " Hrs " + miDiff + " Mins " + secDiff + " Sec"	
}
catch(ex)
{
      return "";
}
	
}
//Returns the Unique file name using current date and time 
function fn_getuniquefilename()
{
	var indate,inthr,intmi,intsec;
	indate = aqConvert.DateTimeToFormatStr(aqDateTime.Today(),"%b_%d_%y");
	inthr = aqDateTime.GetHours(aqDateTime.Now());
	intmi = aqDateTime.GetMinutes(aqDateTime.Now());
	intsec = aqDateTime.GetSeconds(aqDateTime.Now());
	return indate + "-" + inthr + "_" + intmi + "_" + intsec;
	
} 
//Returns the specified format of theinputed date : Format - Sep-20-2016  
function fn_formatdate(dtInput)
{
	var tempDate;
	tempDate = aqConvert.DateTimeToFormatStr(dtInput,"%b-%d-%Y");
	return tempDate;
}
//Returns the specified format of theinputed dateand time : Format - Sep-20-2016 12:30:45 AM
function fn_formatdatetime(dtInput)
{
	var tempDate;
	tempDate = aqConvert.DateTimeToFormatStr(dtInput,"%b-%d-%Y %I:%M:%S %p (%z)");
	return tempDate;
}

//Returns Project Suite Name
function fn_getprojectSuiteName(){
  //Project Suite FileName : C:\Users\akaur\Downloads\Sample_ProjectForHTMLReports\Sample_ProjectForHTMLReports.pjs
  var pjfileName = ProjectSuite.FileName.substring(ProjectSuite.FileName.lastIndexOf('\\')+1).split('.');
  return pjfileName[0];

}
//Returns Project  Name
function fn_getprojectName(){
 // Log.Message("Project FileName : " + Project.FileName);
 // Project FileName : C:\Users\akaur\Downloads\Sample_ProjectForHTMLReports\Sample_ProjectForHTMLReports\Sample_ProjectForHTMLReports_Prj.mds
  var pfileName = Project.FileName.substring(Project.FileName.lastIndexOf('\\')+1).split('.');
  return pfileName[0];
  
}
//**********************************************************************************
// Common Browser functions
//**********************************************************************************
//Function to close all browsers - need improvments
function closeBrowser(){
  var arr = Array("iexplore", "chrome");
  for(var i in arr)  {
    while (Sys.WaitBrowser(arr[i], 1000).Exists){
      var b = Sys.Browser(arr[i]);
      if (b.Exists && b.ChildCount>0 ){
        var page = b.Page("*");
      if (page.Exists) page.Close();  
      }
    }
  }
   if (Sys.WaitProcess("firefox").Exists) Sys.Process("firefox").Close();
}
    
//Function to clear Cookies
function clearBrowserCookies(){
  var objShell = getActiveXObject("WScript.Shell");
  objShell.Run("powershell RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2");
}


//Function to launch new IE browser
function openNewIEBrowser(url){
  openNewBrowser(btIExplorer, url);
  }
//Function to launch new Chrome browser
function openNewCromeBrowser(url){
  openNewBrowser(btChrome, url);
}
//Function to launch new Firefox browser
function openNewFirefoxBrowser(url){
  openNewBrowser(btFirefox, url);
}
//Function to launch new browser
function openNewBrowser(browserStr, url){
  if (url == undefined) {  
    Browsers.Item(browserStr).Run();
  }else{
    Browsers.Item(browserStr).Run(url);
  }
}