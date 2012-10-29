function Launch()
{
  Log.Message("Enter Launch");
  
  TestedApps.Presenter.Run();
  
  Log.Message("click button to launch PowerPoint");
  Aliases.Presenter.PresenterForm.buttonLaunchPpt.WaitProperty("Exists", "true", 10000);  
  Aliases.Presenter.PresenterForm.buttonLaunchPpt.ClickButton();
  
  Log.Message("click continue on activation window");
  Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.WaitProperty("Exists", "true", 100000);
  Aliases.ap6mn.WizardDialog.af10_Introduction.buttonContinue.Click();

  Log.Message("Exit Launch");
}

function ClickArticulateTab()
{
  Log.Message("Enter ClickArticulateTab");
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(547, 46);  
  Log.Message("Exit ClickArticulateTab");
}



function SaveAs(filename)
{
  Log.Message("Enter SaveAs");
  Log.Message("using default output dir: " + Project.Variables.OutputPath);
  Log.Message("Click File|SaveAs");
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.SetFocus();
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(59,20);

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
  
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.SetFocus();
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(59,20);
  
  Log.Message("File | Save");
}

function FileExit()
{
  Log.Message("Enter FileExit");
 
  Aliases.POWERPNT.wndPP12FrameClass.Close();

  Log.Message("Exit FileExit");
}

function Test1()
{
  Aliases.POWERPNT.wndPP12FrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(28, 38);
  Aliases.POWERPNT.wndNetUIToolWindowLayered.NetUIHWND.Click(50, 171);
  Aliases.POWERPNT.wndNetUIToolWindowLayered.NetUIHWND.DblClick(50, 145);
  Aliases.POWERPNT.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.Edit.DblClick(52, 8);
  Aliases.POWERPNT.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.SetText("testpptx");
}