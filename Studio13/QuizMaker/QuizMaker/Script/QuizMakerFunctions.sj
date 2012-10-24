function LaunchQM()
{
  Log.Message("Enter LaunchQM");
  
  //launch QM
  TestedApps.Quizmaker.Run();
  //Clicks at point (108, 27) of the 'buttonContinue' object.
  Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonActivate.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonActivate.Click();
  Aliases.Quizmaker.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.Click();
  Aliases.Quizmaker.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.SetText(ProjectSuite.Variables.StudioActivationKey);
  Aliases.Quizmaker.WizardDialog.nextButton.Click();
  Aliases.Quizmaker.ActivationSuccess.btnLater.Wait("Exists", "true", "10000");
  Aliases.Quizmaker.ActivationSuccess.btnLater.Click();
  
  //create new project
  Aliases.Quizmaker.exMainForm.ucStartPage.LinkLabelPanel.Click();

  Log.Message("Exit LaunchQM");
}

function InsertDummySlide()
{
  Log.Message("Enter InsertDummySlide");
  
  Log.Message("Click ribbon to insert graded question");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(37, 32);
  
  Log.Message("Insert TF from tab conrol")
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click(36, 39);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("Dummy Slide");

  Log.Message("set answer choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("Dummy Answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("Dummy Answer");

  Log.Message("Exit InsertDummySlide");
}

function InsertTF()
{
  Log.Message("Enter InsertTF");
  
  Log.Message("insert graded question from ribbon");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(90, 40);
    
  Log.Message("Insert TF from tab conrol")
 // Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click(36, 39);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("Test String");

  Log.Message("set answer choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("Test String");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("Test String");
  
  Log.Message("Exist InsertTF");
}


function QMSaveAndExit(filename)
{ 
  Log.Message("Enter QMSaveAndExit");
  
  //closing exMainForm generates a save dialogue
  //note: timeing issue in the transition from publish dialogue closing in previus test step
  //      added a wait and clicking to refocus on window
  Aliases.Quizmaker.exMainForm.WaitProperty("Exists", "true", 1000);
  Aliases.Quizmaker.exMainForm.SetFocus();
  Aliases.Quizmaker.exMainForm.Click();
  Aliases.Quizmaker.exMainForm.Close();
  Aliases.Quizmaker.dlgArticulateQuizmaker13.btnSave.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.dlgArticulateQuizmaker13.btnSave.Click();
  
  //set filename and save
  Aliases.Quizmaker.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.Edit.Click();
  Aliases.Quizmaker.dlgSaveAs.DUIViewWndClassName.DirectUIHWND.FloatNotifySink.ComboBox.Edit.SetText(Project.Variables.OutputPath + "\\" + filename);
  Aliases.Quizmaker.dlgSaveAs.btnSave.Click();
    
  
  //terminate application`
  Log.Message("Terminating QM process");
  
  //do we have a timing issue here? does this close call stop the file save operation?
  Aliases.Quizmaker.Close();
  
  Log.Message("Exit QMSaveAndExist");
}

function InsertMC()
{
  //Clicks at point (84, 24) of the 'RibbonTab' object.
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  //Delays the script execution until the specified property equals the specified value or until the specified time limit elapses.
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  //Clicks at point (, ) of the 'ImportChildControl1' object.
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click("", "");
  //Clicks the 'buttonInsert' button.
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  //Clicks at point (, ) of the 'textBoxQuestion' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click("", "");
  //Enters 'testing123' in the 'textBoxQuestion' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing123");
  //Clicks at point (226, 19) of the 'Row_0' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click(226, 19);
  //Simulates a left-button single click in a window or control as specified (relative position, shift keys).
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click("", "");
  //Enters 'testing1234' in the 'ConstantFontTextBox' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("testing1234");
  //Clicks at point (161, 16) of the 'Row_1' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click(161, 16);
  //Simulates a left-button single click in a window or control as specified (relative position, shift keys).
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click("", "");
  //Enters 'testing123' in the 'ConstantFontTextBox' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("testing123");
  //Clicks at point (157, 30) of the 'Row_2' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click(157, 30);
  //Simulates a left-button single click in a window or control as specified (relative position, shift keys).
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click("", "");
  //Enters 'test12' in the 'ConstantFontTextBox' object.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("test12");
  //Clicks the 'RadioButtonEx' button.
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();
}