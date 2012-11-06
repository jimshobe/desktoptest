function ShutdownAndSave(filename)
{
  Log.Message("Enter ShutdownAndSave");

  //Closes the 'MainForm' window.
  Aliases.Engage.MainForm.Close();

  //closing the mainform creates a do you want to save popup
  Log.Message("Click save button");
  Aliases.Engage.dlgArticulateEngage13.btnSave.WaitProperty("Exists", true, 10000);
  Aliases.Engage.dlgArticulateEngage13.btnSave.ClickButton();

  path = Project.Variables.OutputPath + "\\" + filename;
  Log.Message("Saving interaciton as: " + path);
  Aliases.Engage.dlgSaveAs.btnSave.WaitProperty("Exists", true, 10000);
  Aliases.Engage.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.SetText(path);
  Aliases.Engage.dlgSaveAs.btnSave.ClickButton();
  
  //Posts an information message to the test log.
  Log.Message("Exit ShutdownAndSave");
}

function LaunchEngage()
{
  Log.Message("Enter LaunchEngage");
  
  Log.Message("run the app");
  TestedApps.Engage.Run(1, true);
  
  //this statement is true if the activation page is displayed
  if(Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.WaitProperty("Exists", "false", 10000))
  {
    Log.Message("application is not activated");
    
    //continue if in trial period or activate if not
    if(Aliases.Engage.WizardDialog.af10_Introduction.buttonContinue.Exists)
    {
      Log.Message("In trial period so click continue");
      Aliases.Engage.WizardDialog.af10_Introduction.buttonContinue.Click();
    }
    if(Aliases.Engage.WizardDialog.af10_Introduction.buttonActivate.Exists)
    {
      Log.Message("activate application");

      Aliases.Engage.WizardDialog.af10_Introduction.buttonActivate.Click();
      Aliases.Engage.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.SetText(ProjectSuite.Variables.StudioActivationKey);
      
      Aliases.Engage.WizardDialog.nextButton.WaitProperty("Exists", "true", 10000);
      Aliases.Engage.WizardDialog.nextButton.ClickButton();

      Aliases.Engage.ActivationSuccess.btnLater.WaitProperty("Exists", "true", 20000);
      Aliases.Engage.ActivationSuccess.btnLater.ClickButton();
      
      Log.Message("activated successfully");      
    }
  }

  Log.Message("Click create new project link", "");
  Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.WaitProperty("Exists", "true", 10000);
  Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.Click();
 
  Log.Message("Exit LaunchEngage", "");
}

