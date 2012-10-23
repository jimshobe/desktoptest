function Launch()
{
  Log.Message("Enter Launch");
  
  TestedApps.Presenter.Run();
  
  Log.Message("click button to launch PowerPoint");
  Aliases.Presenter.PresenterForm.buttonLaunchPpt.WaitProperty("Exists", "true", 10000);  
  Aliases.Presenter.PresenterForm.buttonLaunchPpt.ClickButton();
  
  Log.Message("click continue on activation window");
  Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.WaitProperty("Exists", "true", 10000);
  Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.Click();

  Log.Message("Exit Launch");
}

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