function VerifyDummySlide()
{
  Log.Message("Enter VerifyDummySlide");
  Aliases.iexplore.wndIEFrame.Position(50,50,1350,950);
  var dummySlideCoordinates = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), "DummySlideWebOutput", 0, 0, false, false, 20);
  if(dummySlideCoordinates == null)
  {
    Log.Warning("Dummy slide isn't the same.");
  }
  //click on the answer
  var dummyAnswer = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), "DummyAnswer1", 0, 0, false, false, 20);  
  if(dummyAnswer == null)
  {
    Log.Warning("Dummy answer can't be found.");
  }
  else
  {
    x = dummyAnswer.left + (dummyAnswer.Right - dummyAnswer.Left)/2;
    y = dummyAnswer.Top + (dummyAnswer.Bottom - dummyAnswer.Top)/2;
    Aliases.iexplore.wndIEFrame.Click(x, y);
  }
  //click submit
  var submitButton = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), "SubmitButton", 0, 0, false, false, 20);  
  if(submitButton == null)
  {
    Log.Warning("Submit button can't be found.");
  }
  else
  {
    x = submitButton.left + (submitButton.Right - submitButton.Left)/2;
    y = submitButton.Top + (submitButton.Bottom - submitButton.Top)/2;
    Aliases.iexplore.wndIEFrame.Click(x, y);   
  }
  
  //make sure the correct feedback is shown 
  var correctDummy = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), "CorrectDummyFeedback", 0, 0, false, false, 20);  
  if(correctDummy == null)
  {
    Log.Warning("Correct Dummy feedback can't be found.");
  }

  //click continue
  var continueButton = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), "ContinueButton", 0, 0, false, false, 20);  
  if(continueButton == null)
  {
    Log.Warning("Continue button can't be found.");
  }
  else
  {
    x = continueButton.left + (continueButton.Right - continueButton.Left)/2;
    y = continueButton.Top + (continueButton.Bottom - continueButton.Top)/2;
    Aliases.iexplore.wndIEFrame.Click(x, y);   
  }
  
  Log.Message("Exit VerifyDummySlide");
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
  
  Log.Message("Exit InsertTF");
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
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing multiple choice");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();
  
  Log.Message("Exit InsertMultipleChoice");
}


function InsertMatchingDragDrop()
{
  Log.Message("Insert Matching Drag and Drop question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing matching drag and drop");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Enter Matches text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("match a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("match b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("match c");

  Log.Message("Exit InsertMatchingDragDrop");
}


function InsertWordBank()
{
  Log.Message("Insert Word Bank question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing word bank");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.RadioButtonEx.ClickButton();

  Log.Message("Exit InsertWordBank");
}

//Insert survey questions
function InsertShortAnswer()
{
  Log.Message("Insert Short Answer question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing short answer");

  Log.Message("Exit InsertShortAnswer");
}


function InsertRankingDropDown()
{
  Log.Message("Insert Ranking Drop-down question from Insert Questions dialog box");  
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl7.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl7.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing ranking drop-down");
 
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Exit InsertRankingDropDown");
}


function InsertHowMany()
{
  Log.Message("Insert How Many question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl8.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl8.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing how many");

  Log.Message("Exit InsertHowMany");
}


function InsertFreeformPickOne()
{
  Log.Message("Insert Pick One freeform question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click();
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

  Log.Message("Exit InsertFreeFormPickOne");
}

function InsertFreeformHotspot()
{
  Log.Message("Insert Hotspot freeform question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl4.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Insert an image");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonImage.Click();
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.WaitProperty("Exists", true, 10000);
  var testFile = ProjectSuite.Variables.TestDataDir+ "\\Tulips.jpg";
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.SetText(testFile);
  Aliases.Quizmaker.dlgInsertPicture.btnOpen.Click();
  
  Log.Message("Insert a hotspot");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.PopupMenu.Click("[1]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.hotspotPanel.Drag(91, 27, 71, 76);  

  Log.Message("Exit InsertFreeformHotspot");  
}


function InsertMultipleResponse()
{
  Log.Message("Insert Multiple Choice question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();

  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing multiple response");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("choice d");

  Log.Message("Set Choice A and Choice D as correct answers");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.CheckBoxEx.ClickButton();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.CheckBoxEx.ClickButton();

  Log.Message("Exit InsertMultipleResponse");  
}


function InsertFillInTheBlank()
{
  Log.Message("Insert Fill in the Blank question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing fill in the blank");
  
  Log.Message("Enter Acceptable Answers");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  
  Log.Message("Exit InsertFillInTheBlank");    
}

function InsertMatchingDropDown()
{
  Log.Message("Insert Matching Drop-down question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing matching drop-down");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Enter Matches text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("match a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("match b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("match c");

  Log.Message("Exit InsertMatchingDropDown");  
}


function InsertSequenceDragAndDrop()
{
  Log.Message("Insert Sequence Drag and Drop question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl7.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl7.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing sequence drag and drop");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Exit InsertSequenceDragAndDrop"); 
}


function InsertSequenceDropDown()
{
  Log.Message("Insert Sequence Drop-down question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl8.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl8.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing sequence drop-down");
  
  Log.Message("Enter Choices text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("choice a");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("choice b");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("choice c");

  Log.Message("Exit InsertSequenceDropDown"); 
}


function InsertNumeric()
{
  Log.Message("Insert Numeric question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl9.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl9.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing numeric");
  
  Log.Message("Enter acceptable values");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.NumericInteractionNumericControl.groupBoxMain.numericControl.comboBoxValue1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.NumericInteractionNumericControl.groupBoxMain.numericControl.comboBoxValue1.ClickItem('EqualTo');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.NumericInteractionNumericControl.groupBoxMain.numericControl.spinButtonA1.SpinTextBox.DblClick();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelMainContainer.panelMain.NumericInteractionNumericControl.groupBoxMain.numericControl.spinButtonA1.SpinTextBox.SetText("12");
  
  Log.Message("Exit InsertNumeric"); 
}

function InsertGradedHotspot()
{
  Log.Message("Insert Hotspot graded question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(84, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl10.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.gradedPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl10.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("Testing hotspot graded question");
  
  Log.Message("Insert an image");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonImage.Click();
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.WaitProperty("Exists", true, 10000);
  var testFile = ProjectSuite.Variables.TestDataDir+ "\\Tulips.jpg";
  Aliases.Quizmaker.dlgInsertPicture.ComboBoxEx32.Keys(testFile);
  Aliases.Quizmaker.dlgInsertPicture.btnOpen.Click();
  
  Log.Message("Insert a hotspot");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.panelBottom.buttonHotspot.PopupMenu.Click("[1]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.HotspotInteractionControl.panelMainContainer.panelLeftOfMain.HotspotInteractionHotspotControl.panelMain.hotspotPanel.Drag(93, 43, 69, 67);  

  Log.Message("Exit InsertGradedHotspot");  
}


function InsertLikertScale()
{
  Log.Message("Insert Likert Scale question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing likert scale");

  Log.Message("Enter statements");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("statement 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("statement 2");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("statement 3");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.LikertQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("statement 4");  
  
  Log.Message("Exit InsertLikertScale");
}


function InsertPickOne()
{
  Log.Message("Insert Pick One question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl1.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing pick one");

  Log.Message("Enter choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("statement 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("statement 2");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("statement 3");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("statement 4");  
  
  Log.Message("Exit InsertPickOne");
}


function InsertPickMany()
{
  Log.Message("Insert Pick Many question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing pick many");

  Log.Message("Enter choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("statement 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("statement 2");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("statement 3");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("statement 4");  
  
  Log.Message("Exit InsertPickMany");
}


function InsertWhichWord()
{
  Log.Message("Insert Which Word question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing which word");

  Log.Message("Enter choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("statement 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("statement 2");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("statement 3");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("statement 4");  
  
  Log.Message("Exit InsertWhichWord");
}


function InsertEssay()
{
  Log.Message("Insert Essay question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl5.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FillInBlankInteractionControl.panelTopHidden.panelTop.textBoxQuestion.SetText("testing essay");

  Log.Message("Exit InsertEssay");
}


function InsertRankingDragDrop()
{
  Log.Message("Insert Ranking Drag and Drop question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(150, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.surveyPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl6.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Enter Question text");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Select();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelTopHidden.panelTop.textBoxQuestion.SetText("testing ranking drag and drop");

  Log.Message("Enter choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("statement 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("statement 2");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("statement 3");  
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.QuestionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("statement 4");  
  
  Log.Message("Exit InsertRankingDragDrop");
}


function InsertFreeformDragDrop()
{
  Log.Message("Insert Drag and Drop freeform question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Insert 4 shapes");
  Aliases.Quizmaker.exMainForm.Keys("~[ReleaseLast]nsh");
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(11,70);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(134, 99, 217, 134);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(53,107);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(205, 301, 188, 155);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(34,66);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(427, 77, 324, 192);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(109,109);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(475, 297, 265, 216);
  
  Log.Message("Switch to Form View");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.panelRight.CollapsibleControl.SidePanelHost.QuestionSidePanel.panelTop.toggleButtons.Click(65,25);
  
  Log.Message("Set choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.FreeFormComboBox.ClickItem('Rectangle 1');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.FreeFormComboBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.FreeFormComboBox.ClickItem('Rectangle 2');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.FreeFormComboBox.ClickItem('Triangle 1');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.FreeFormComboBox.Keys("[Tab]");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.DragDropInteractionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.FreeFormComboBox.ClickItem('Trapezoid 1');

  Log.Message("Exit InsertFreeFormDragDrop");
}


function InsertFreeformPickMany()
{
  Log.Message("Insert Pick Many freeform question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl2.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
  
  Log.Message("Insert 4 shapes");
  Aliases.Quizmaker.exMainForm.Keys("~[ReleaseLast]nsh");
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(11,70);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(134, 99, 217, 134);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(53,107);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(205, 301, 188, 155);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(9,260);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(427, 77, 324, 192);
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab1.Click(413,25);
  Aliases.Quizmaker.xbc4f5b163cf61008.x22b8fc887c9fa272.Click(93,146);
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.panelMiddle.panelSlide.stage.Drag(475, 297, 265, 216);
  
  Log.Message("Switch to Form View");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.panelRight.CollapsibleControl.SidePanelHost.QuestionSidePanel.panelTop.toggleButtons.Click(65,25);
  
  Log.Message("Set choices");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.FreeFormComboBox.ClickItem('Rectangle 1');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.FreeFormComboBox.ClickItem('Triangle 1');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.FreeFormComboBox.ClickItem('Star 1 - \"Explosion 1 1\"');
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.FreeFormComboBox.ClickItem('Smiley Face 1');

  Log.Message("Set Choice A as correct answer");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.CheckBoxEx.ClickButton();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormInteractionControlBase.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.CheckBoxEx.ClickButton();

  Log.Message("Exit InsertFreeFormPickOne");
}


function InsertTextEntry()
{
  Log.Message("Insert Text Entry question from Insert Questions dialog box");
  Aliases.Quizmaker.exMainForm.CustomRibbon.RibbonTab.Click(210, 24);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.WaitProperty("Exists", true, 10000);
  Aliases.Quizmaker.InsertSlidesTabDialog.panelContainer.QuizzingPanel.panelContainer.freeformPanel.panelMain.slidePicker.ImportGroupControl.ImportChildControl3.Click();
  Aliases.Quizmaker.InsertSlidesTabDialog.panelButtons.InsertButtons.buttonInsert.ClickButton();
    
  Log.Message("Switch to Form View");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.panelRight.CollapsibleControl.SidePanelHost.QuestionSidePanel.panelTop.toggleButtons.Click(65,25);
  
  Log.Message("Set acceptable answers");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_0.ConstantFontTextBox.SetText("answer 1");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_1.ConstantFontTextBox.SetText("answer 2");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_2.ConstantFontTextBox.SetText("answer 3");
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.Click();
  Aliases.Quizmaker.exMainForm.panelProject.tabControl.CustomTabPage.exSlideView.BaseUserControl.FreeFormTextEntryQuestionControl.panelMainContainer.panelMain.gridListChoices.ListPanel.GridListInternal.Row_3.ConstantFontTextBox.SetText("answer 4");
  
  Log.Message("Exit InsertTextEntry");
}