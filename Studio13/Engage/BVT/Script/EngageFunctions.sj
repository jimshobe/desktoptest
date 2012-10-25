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