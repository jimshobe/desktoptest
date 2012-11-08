//USEUNIT Environment

function GeneralEvents_OnLogError(Sender, LogParams)
{
  if(true == Project.Variables.Pass)
  {
    Log.Message("Enter Event_OnLogError");
  
    Log.Message("Set Pass = false");
    Project.Variables.Pass = false;
  
    Log.Message("Exit Event_OnLogError");  
  }
}

function GeneralEvents_OnStopTest(Sender)
{
  //if there has been a failure lets save results and bail out
  if(false == Project.Variables.Pass)
  {
    Log.Message("Test failure so save results and exit");
    Environment.SaveResults();
  }  
}