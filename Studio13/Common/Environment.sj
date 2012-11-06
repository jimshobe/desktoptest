//TODO: add error handling in here

//the install scripts create a folder under path\instll with the build
//version number and time stamp.  We read that here to determine the
//product version we have installed
function GetInstalledBuild(path)
{
  Log.Message("Enter GetInstalledBuildPath");
  
  var objFSO = new ActiveXObject("Scripting.FileSystemObject");
  var objFolder = objFSO.GetFolder(path);
  var colSubFolder = new Enumerator(objFolder.SubFolders); 
  
  var installFolderName;
  colSubFolder.moveFirst();
  var installFolder = colSubFolder.item();
  var folderName = installFolder.Name; 
 
  Log.Message("returning " + folderName);
  return folderName;
}

function InitTestEnvironment(testproject)
{
  Log.Message("Enter InitTestEnvironment");
  
  //instantiate config object - assumes config file in projectsuite dir
  Log.Message("instantiate tconfig object");
  var s = ProjectSuite.ConfigPath + "\\tconfig.xml";
  var objConfig = dotNET.TConfig.TConfig.zctor(s);
  
  Log.Message("get restults output path");
  Project.Variables.OutputPath = objConfig.ResultsOutputPath(testproject).OleValue;
  Log.Message("get activationkey");
  ProjectSuite.Variables.StudioActivationKey = objConfig.ActivationKey().OleValue;
  Log.Message("get testdatadir");
  ProjectSuite.Variables.TestDataDir = objConfig.TestDataDir().OleValue;
 
  Log.Message("Output path set to: " + Project.Variables.OutputPath);
  Log.Message("Activation code: " + ProjectSuite.Variables.StudioActivationKey);
  Log.Message("TestDataDir set to: " + ProjectSuite.Variables.TestDataDir);

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

function PackResults()
{
  var WorkDir, FileName, FileList, ArchivePath;
  WorkDir = Project.ConfigPath + "Log\\ExportedResults\\";
  FileName = WorkDir + "MyFile.mht";
  Log.SaveResultsAs(FileName, 2);
  FileList = slPacker.GetFileListFromFolder(WorkDir);
  ArchivePath = WorkDir + "CompressedResults";
  if (slPacker.Pack(FileList, WorkDir, ArchivePath))
    Log.Message("Files compressed successfully.");
}