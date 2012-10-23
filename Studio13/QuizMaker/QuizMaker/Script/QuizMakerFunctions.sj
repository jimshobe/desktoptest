function LaunchQM()
{
  Log.Message("Enter LaunchQM");
  
  //launch QM
  TestedApps.Quizmaker.Run();
  //Clicks at point (108, 27) of the 'buttonContinue' object.
  Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonContinue.WaitProperty("Exists", "true", 10000);
  Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonContinue.Click();
  
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