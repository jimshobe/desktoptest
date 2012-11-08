//TODO: add error handling in here

function InitTestEnvironment(testproject)
{
  Log.Message("Enter InitTestEnvironment");
  
  //set pass value to true. this gets set to false on log.error
  Project.Variables.Pass = true;
  
  Project.Variables.ProjectName = testproject;
  
  //instantiate config object - assumes config file in projectsuite dir
  Log.Message("instantiate tconfig object");
  var s = ProjectSuite.ConfigPath + "\\tconfig.xml";
  var objConfig = dotNET.TConfig.TConfig.zctor(s);
  
  Log.Message("get restults output path");
  Project.Variables.OutputPath = objConfig.ResultsOutputPath(Project.Variables.ProjectName).OleValue;
  Project.Variables.ResultsUrl = objConfig.ResultsWebUrl(Project.Variables.ProjectName).OleValue;
  Log.Message("get activationkey");
  ProjectSuite.Variables.StudioActivationKey = objConfig.ActivationKey().OleValue;
  Log.Message("get testdatadir");
  ProjectSuite.Variables.TestDataDir = objConfig.TestDataDir().OleValue;
  Log.Message("get installed build number");
  ProjectSuite.Variables.InstalledBuildNumber = objConfig.InstalledBuildNumber().OleValue;
  
  Log.Message("get email info");
  if(objConfig.SendMail())
  {
    Project.Variables.SendMail = true;
    Project.Variables.SendMailTo = objConfig.SendMailTo().OleValue;
    
    var s = Project.Variables.SendMailTo;
    Log.Message("Send email to:" + Project.Variables.SendMailTo);
  }
  else
    Log.Message("do not email results");
 
  Log.Message("Output path set to: " + Project.Variables.OutputPath);
  Log.Message("Result Url set to: " + Project.Variables.ResultsUrl);
  Log.Message("Activation code: " + ProjectSuite.Variables.StudioActivationKey);
  Log.Message("TestDataDir set to: " + ProjectSuite.Variables.TestDataDir);
  Log.Message("InstalledBuildNumber set to: " + ProjectSuite.Variables.InstalledBuildNumber);

  Log.Message("Exit InitTestEnvironment");
}


function CopyResultsFile()
{
  Log.Message("Enter CopyResultsFile");
  
  Log.Message("Config Dir = " + Project.ConfigPath);
  var source = Project.ConfigPath + "results.mht";
  var destination = Project.Variables.OutputPath + "\\results.mht";
  
  Log.Message("source: " + source);
  Log.Message("destination: " + destination);
  
  //copy file from config dir to output dir
  var objFSO = new ActiveXObject("Scripting.FileSystemObject");
  objFSO.CopyFile(source, destination);
  
  //delete results file from source
  var f = objFSO.GetFile(source);
  f.Delete();
  
  Log.Message("Exit CopyResultsFile");
}

function SaveResults()
{
  Log.Message("Enter SaveResults");
    Log.Message(Project.FileName);
  var path = Project.Variables.OutputPath + "\\LogFiles";
  
  Log.Message("Create results folder: " + path);
  var objFSO = new ActiveXObject("Scripting.FileSystemObject");
  if(!objFSO.FolderExists(path))
    objFSO.CreateFolder(path);
  
  Log.Message("save results as HTML");
  Log.SaveResultsAs(path, 1); 
  
  SendResultsEmail();
  
  Log.Message("Exit SaveResults");
}

function SendResultsEmail()
{
  try
  {
    Log.Message("Enter SendResultsMail");
  
    var sSubject = ProjectSuite.Variables.InstalledBuildNumber + " ";
    sSubject += Project.Variables.ProjectName + " ";
    sSubject += "TestResults: ";
    if(true == Project.Variables.Pass)
      sSubject += "PASS";
    else
      sSubject += "FAIL";
    Log.Message("subject: " + sSubject); 

    var sBody = "Full test log: " + Project.Variables.ResultsUrl + "\\LogFiles\\index.html";
    Log.Message("Body: " + sBody);  
  
    var objMessage = Sys.OleObject("CDO.Message");
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "qa-server"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objMessage.Configuration.Fields.Update();

    objMessage.From = "jshobe@articulate.com";
    objMessage.To = Project.Variables.SendMailTo;
    objMessage.Subject = sSubject;
    objMessage.HTMLBody = sBody;
    objMessage.Send();
  }
  catch (exception)
  {
    Log.Error("E-mail cannot be sent", exception.description);
    return false;
    Log.Message("Exist SendEmail");
  }
 
  Log.Message("Message to <" + mTo + "> was successfully sent");
  return true;
  
  Log.Message("Exit SendEmail");
}
