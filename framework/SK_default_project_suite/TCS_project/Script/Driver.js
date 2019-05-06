//USEUNIT CustomFunctions
const STEP_PASSED = 1;
const STEP_FAILED = 0;

function Unit3()
{      //create report and table and header
//if want to create report at different path 
//CustomFunctions.setLogsPath("C:\\AutomationLogsTest\\");

      CustomFunctions.fn_createreportfile("SYS");
      
      //Run project 1
      CustomFunctions.fn_createprojectnamerow(CustomFunctions.fn_getprojectName());
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);
      CustomFunctions.fn_createteststep(STEP_FAILED,"One line text on Action","Expe Result","Actual Result",true);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);

      //Run project 2
      CustomFunctions.fn_createprojectnamerow(CustomFunctions.fn_getprojectName());
      CustomFunctions.fn_createteststep(STEP_FAILED,"One line text on Action","Expe Result","Actual Result",true);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);
      CustomFunctions.fn_createteststep(STEP_FAILED,"One line text on Action","Expe Result","Actual Result",true);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);
      CustomFunctions.fn_createteststep(STEP_PASSED,"One line text on Action","Expe Result","Actual Result",false);

      //Close table and html
      CustomFunctions.fn_generatehighlevelreport("SYS");

      }
      
