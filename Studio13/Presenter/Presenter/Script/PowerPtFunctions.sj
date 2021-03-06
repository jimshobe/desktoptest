function Launch()
{
  Log.Message("Enter Launch");

  try
  {  
    TestedApps.Presenter.Run();
  
    Log.Message("click button to launch PowerPoint");
    Aliases.Presenter.PresenterForm.buttonLaunchPpt.WaitProperty("Exists", "true", 10000);  
    Aliases.Presenter.PresenterForm.buttonLaunchPpt.ClickButton();
  
    if(Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.Exists)
    {
      Log.Message("In trial period.  Click continue");
      Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.Click();
    }
    else
    {
      Log.Message("Trial has expired so might have to activate");
      
      ClickArticulateTab();
      
      Log.Message("Insert character to invoke activation window");
      Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(391, 107);
      Aliases.POWERPNT.wndNetUIToolWindow.NetUIHWND.Click(35, 6);
  
      //if already acivated just choose to not save and move on
      if(Aliases.POWERPNT.dlgArticulatePresenter.btnNo.WaitProperty("Exists", "true", 10000))
      {
        Log.Message("already activated so click not to save and move on");
        Aliases.POWERPNT.dlgArticulatePresenter.btnNo.ClickButton();
      }
      else
      {
        Log.Message("must activate");
        
        //Clicks at point (115, 27) of the 'WindowsForms10Window8app033c0d9d' object.
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d.WindowsForms10Window8app033c0d9d.WindowsForms10Window8app033c0d9d.WaitProperty("Exists", "true", 10000);
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d.WindowsForms10Window8app033c0d9d.WindowsForms10Window8app033c0d9d.Click();
    
        //Clicks at point (209, 11) of the 'Item' object.
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d.WindowsForms10Window8app033c0d9d1.Item.Click(209, 11);
      
        Log.Message("Click activate button");
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d.btnActivate.ClickButton();
    
        Log.Message("Register Later");
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d1.btnRegisterLater.WaitProperty("Exists", "True", 20000);
        Aliases.ap6mn.wndWindowsForms10Window8app033c0d9d1.btnRegisterLater.ClickButton();
    
        Log.Message("Click no to not save");
        Aliases.POWERPNT.dlgArticulatePresenter.btnNo.ClickButton();
      }
    }
    //if 
    //if(Aliases.ap6mn.WizardDialog.af10_Introduction.
      
  }
  catch (e)
  {
    Project.Variables.Pass = false;
    Log.Error(e.descripton);
  }
  finally
  {
    Log.Message("Exit Launch");
  }
}


/*
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
*/






function ClickArticulateTab()
{
  Log.Message("Enter ClickArticulateTab");
  
  //Clicks at point (652, 43) of the 'NetUIHWND' object.
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(652, 43);
  
  Log.Message("Exit ClickArticulateTab");
}



function SaveAs(filename)
{
  Log.Message("Enter SaveAs");
  Log.Message("using default output dir: " + Project.Variables.OutputPath);
  Log.Message("Click File|SaveAs");
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.SetFocus();
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(20, 39);
  Aliases.POWERPNT.wndPPTFrameClass.FullpageUIHost.NetUIHWND.WaitProperty("Exists", "true", 10000);
  Aliases.POWERPNT.wndPPTFrameClass.FullpageUIHost.NetUIHWND.Click(34, 104);
  

  Log.Message("Enter filename text and click save");
  Aliases.POWERPNT.dlgSaveAs.btnSave.WaitProperty("Exists", "true", 10000);
  Aliases.POWERPNT.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.Edit.SetFocus();
  Aliases.POWERPNT.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.Click();
  Aliases.POWERPNT.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.SetText(Project.Variables.OutputPath + filename);
  
  Aliases.POWERPNT.dlgSaveAs.btnSave.ClickButton();
  
  Log.Message("Exit SaveAs");
}

function FileSave()
{
  Log.Message("Enter File|Save");
  
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(26, 43);
  Aliases.POWERPNT.wndPPTFrameClass.FullpageUIHost.NetUIHWND.Click(35, 80);

  Log.Message("File | Save");
}

function FileExit()
{
  Log.Message("Enter FileExit");
  
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(24, 43);
  Aliases.POWERPNT.wndPPTFrameClass.FullpageUIHost.NetUIHWND.Click(54, 445);

  Log.Message("Exit FileExit");
}
