function LaunchQM()
{
  Log.Message("Enter LaunchQM");
  
  //launch QM
  TestedApps.Quizmaker.Run();

  //Activate QM
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

function LaunchQMAlways()
{
  Log.Message("Enter LaunchQMAlways");
  
  //launch QM
  TestedApps.Quizmaker.Run();
  
  //activate it if it hasn't already been activated
  if(Aliases.Quizmaker.WizardDialog.af10_Introduction.Exists==true)
  {
    Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonActivate.WaitProperty("Exists", "true", 10000);
    Aliases.Quizmaker.WizardDialog.af10_Introduction.buttonActivate.Click();
    Aliases.Quizmaker.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.Click();
    Aliases.Quizmaker.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.SetText(ProjectSuite.Variables.StudioActivationKey);
    Aliases.Quizmaker.WizardDialog.nextButton.Click();
    Aliases.Quizmaker.ActivationSuccess.btnLater.Wait("Exists", "true", "10000");
    Aliases.Quizmaker.ActivationSuccess.btnLater.Click();
  }
  
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
//  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
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

function InsertMultipleChoice()
{
  Log.Message("Insert Multiple Choice Question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing multiple choice");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();
}


function InsertMatchingDragDrop()
{
  Log.Message("Insert Matching Drag and Drop question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing matching drag and drop");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");

  Log.Message("Enter Matches text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("match a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("match b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("match c");

}


function InsertWordBank()
{
  Log.Message("Insert Word Bank question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing word bank");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();
}

//Insert survey questions
function InsertShortAnswer()
{
  Log.Message("Insert Short Answer question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Keys("testing short answer");
}


function InsertRankingDropDown()
{
  Log.Message("Insert Ranking Drop-down question from Insert Questions dialog box");  
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing ranking drop-down");
 
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");
}


function InsertHowMany()
{
  Log.Message("Insert How Many question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Keys("testing how many");
}


function InsertPickOne()
{
  Log.Message("Insert PickOne question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Insert 2 shapes");
  Aliases.Quizmaker.exMainForm.Keys("~[ReleaseLast]nsh");
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(11,70);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(134, 99, 217, 134);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(53,107);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(205, 301, 188, 155);
  
  Log.Message("Switch to Form View");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.panelRight.CollapsibleControl.SidePanelHost.QuestionSidePanel.panelTop.toggleButtons.Click(65,25);
  
  Log.Message("Set choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.FreeFormComboBox.ClickItem('Rectangle 1');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.FreeFormComboBox.ClickItem('Triangle 1');

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();
}

function InsertFreeformHotspot()
{
  Log.Message("Insert Hotspot freeform question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
/*  Log.Message("Insert an image");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonImage.Click();
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.Keys("C:\\test_files\\Tulips.jpg");
  Aliases.Quizmaker.dlgInsertPicture.btnOpen.Click();
*/
    
  Log.Message("Insert a hotspot");

  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.PopupMenu.Click("[1]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.hotspotPanel.Drag(91, 27, 71, 76);  

}


function InsertMultipleResponse()
{
  Log.Message("Insert Multiple Choice question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing multiple response");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.Keys("choice d");

  Log.Message("Set Choice A and Choice D as correct answers");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.CheckBoxEx.ClickButton();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.CheckBoxEx.ClickButton();
  
}


function InsertFillInTheBlank()
{
  Log.Message("Insert Fill in the Blank question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Keys("testing fill in the blank");
  
  Log.Message("Enter Acceptable Answers");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  
}

/*
function InsertMatchingDropDown()
{
  Log.Message("Insert Matching Drop-down question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Keys("testing matching drop-down");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("choice c");

  Log.Message("Enter Matches text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("match a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("match b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("match c");

}
*/