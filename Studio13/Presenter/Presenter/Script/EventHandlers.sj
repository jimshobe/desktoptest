//USEUNIT Environment

function GeneralEvents_OnLogError(Sender, LogParams)
{
  if(true == Project.Variables.Pass)
  {
    //only run this block on the first error to prevent circular loops
    //from subsequent error logging
    Log.Message("Enter Event_OnLogError");
  
    Log.Message("Set Pass = false");
    Project.Variables.Pass = false;
    
    Log.Message("Terminate Application");
    TestedApps.Presenter.Terminate();

    if(Project.Variables.IsBvt)
    {
      Log.Message("Running BVT test so halt execution and publish results");
      Runner.Halt();
    }
  }
  
  Log.Message("stop running test: " + Project.TestItems.Current.Name);
  Runner.Stop(true);
}

function GeneralEvents_OnStopTest(Sender)
{
  //if there has been a failure 
  if(false == Project.Variables.Pass)
  {
    if(Project.Variables.IsBvt)
    { 
      Log.Message("BVT Test failure so save results and exit");
      Environment.SaveResults();
    }
  }  
}