var ResultsOutputPath;

//TODO: add error handling in here

//the install scripts create a folder under path\instll with the build
//version number and time stamp.  We read that here to determine the
//product version we have installed
function GetInstalledBuild(path)
{
  Log.Message("Enter GetInstalledVersion");
  
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

function SetResultsOutputPath(testproject)
{
  Log.Message("Enter SetResultsOutputPath");
 
  var objFSO = new ActiveXObject("Scripting.FileSystemObject");
 
  Log.Message("out path = " + Project.Variables.OutputPath);
   
 // Log.Message("working path: " + Project.Variables.)
  var installScriptsPath = Project.ConfigPath + "install\\";
  var path = Project.ConfigPath + GetInstalledBuild(installScriptsPath);
    
  //if path does not exist, create it
  if(!objFSO.FolderExists(path))
  {
    //use double backslashes for CreateFolder method
    var temp = path.replace(/\\/g,"\\\\");
    Log.Message("create folder " + temp);
    objFSO.CreateFolder(temp);
  }
  
  path = path + "\\" + testproject;
    
  //create results folder if it does not exist
  if(!objFSO.FolderExists(path))
  {
    var temp = path.replace(/\\/g,"\\\\"); 
    Log.Message("Create folder: " + temp);
    objFSO.CreateFolder(temp);
    path = path + "\\";
  }
  else //folder already exists so append a number to make it unique
  {   
    var temp = path;
    var i = 0;
    while (objFSO.FolderExists(temp))
    {
      i++;
      temp = path;
      temp = path + "_" + i;
    }
    path = temp + "\\";
    temp = path.replace(/\\/g, "\\\\");
    Log.Message("Create folder: " + temp); 
    objFSO.CreateFolder(temp);
  }

  //append path with Output
  path = path + "Output" + "\\";
  Log.Message("Create folder: " + temp);
  var temp = path.replace(/\\/g,"\\\\");
  objFSO.CreateFolder(temp);
  
  Project.Variables.OutputPath = path;
  Log.Message("Results Path = " + Project.Variables.OutputPath);
  Log.Message("Exit SetResultsOutputPath");
}

function CopyResultsFile()
{
  Log.Message("Enter CopyResultsFile");
  
  Log.Message("Config Dir = " + Project.ConfigPath);
  var source = Project.ConfigPath + "results.mht";
  
  //copy file from config dir to output dir
  var objFSO = new ActiveXObject("Scripting.FileSystemObject");
  objFSO.CopyFile(source, Project.Variables.OutputPath);
  
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