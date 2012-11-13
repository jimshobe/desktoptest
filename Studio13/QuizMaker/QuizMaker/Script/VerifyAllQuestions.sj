function MatchSlide(KnownGoodSlide)
{
  Log.Message("Enter MatchSlide");
  
  //move and resize the IE window
  Aliases.iexplore.wndIEFrame.SetFocus();
  Aliases.iexplore.wndIEFrame.Position(50,50,1350,950);
  
  //give the page time to load
  
  if(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).WaitProperty("Exists", "true", 50000))
  {
    Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture().SaveToClipboard();
    var time = 0;
    do
    {
      aqUtils.Delay(100);
      time += 100;
      var slideCoordinates = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), KnownGoodSlide, 0, 0, false, false, 80);
    }
    while(slideCoordinates == null && time<10000);
  
    if(slideCoordinates == null)
    {
      Log.Error("Slide isn't the same.");
    }
  }
  else
  {
    Log.Error("Web output never loaded");
  }
  
  Log.Message("Exit MatchSlide");
}


function SelectAnswer(GivenAnswer)
{
  Log.Message("Enter SelectAnswer");

  var theAnswer = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), GivenAnswer, 0, 0, false, false, 20);  
  if(theAnswer == null)
  {
    Log.Error("The answer can't be found.");
  }
  else
  {
    var x = theAnswer.left + (theAnswer.Right - theAnswer.Left)/2;
    var y = theAnswer.Top + (theAnswer.Bottom - theAnswer.Top)/2;
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(x,y);
   }

  Log.Message("Exit SelectAnswer");
}


function DragAnswer(GivenAnswer, DropLocation)
{
  Log.Message("Enter DragAnswer");

  Log.Message("Find the answer and drop location");
  var theAnswer = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), GivenAnswer, 0, 0, false, false, 20);
  var theDropLocation = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), DropLocation, 0, 0, false, false, 10);

  if(theAnswer == null)
  {
    Log.Error("The answer can't be found.");
  }
  else if(theDropLocation == null)
  {
    Log.Error("The drop location can't be found");
  }
  else
  {
    var xAnswer = theAnswer.left + (theAnswer.Right - theAnswer.Left)/2;
    var yAnswer = theAnswer.Top + (theAnswer.Bottom - theAnswer.Top)/2;
    var xDrop = theDropLocation.left + (theDropLocation.Right - theDropLocation.Left)/2;
    var yDrop = theDropLocation.Top + (theDropLocation.Bottom - theDropLocation.Top)/2;
   // Log.Message(xAnswer + ", " + yAnswer + ", " + xDrop + ", " + yDrop);
    //Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(xDrop,yDrop);
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Drag(xAnswer, yAnswer, (xDrop-xAnswer), (yDrop-yAnswer));
   }
   

   
   
   

  Log.Message("Exit SelectAnswer");
}

function ClickSubmit()
{
  Log.Message("Enter ClickSubmit");

  var submitButton = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), "SubmitButton", 0, 0, false, false, 20);  
  if(submitButton == null)
  {
    Log.Error("Submit button can't be found.");
  }
  else
  {
    var x = submitButton.left + (submitButton.Right - submitButton.Left)/2;
    var y = submitButton.Top + (submitButton.Bottom - submitButton.Top)/2;
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(x,y); 
  }

  Log.Message("Exit ClickSubmit");
}


function CheckFeedback(FeedbackImage)
{
  Log.Message("Enter CheckDefaultFeedback");

  var rightFeedback = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), FeedbackImage, 0, 0, false, false, 20);  
  if(rightFeedback == null)
  {
    Log.Error("The right feedback can't be found.");
  }
 
  Log.Message("Exit CheckDefaultFeedback");
}


function ClickContinue()
{
  Log.Message("Enter ClickContinue");
  
  var continueButton = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), "ContinueButton", 0, 0, false, false, 20);  
  if(continueButton == null)
  {
    Log.Error("Continue button can't be found.");
  }
  else
  {
    var x = continueButton.left + (continueButton.Right - continueButton.Left)/2;
    var y = continueButton.Top + (continueButton.Bottom - continueButton.Top)/2;
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(x,y);
    Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).HoverMouse(40,40)
  }
  
  Log.Message("Exit ClickContinue");
}


//Graded questions
function VerifyDummySlide()
{
  Log.Message("Enter VerifyDummySlide");

  //check to make sure the slide matches the known good output
  MatchSlide("DummySlideWebOutput");
 
  //click on the answer
  SelectAnswer("DummyAnswer1");

  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();

  Log.Message("Exit VerifyDummySlide");
}

function VerifyTF()
{
  Log.Message("Enter VerifyTF");
 
 //check to make sure the slide matches the known good output
  MatchSlide("TFSlideWebOutput");
 
  //click on the answer
  SelectAnswer("True");

  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit VerifyTF");
}

function VerifyMultipleChoice()
{
  Log.Message("Insert VerifyMultipleChoice");
 
  //check to make sure the slide matches the known good output
  MatchSlide("MultipleChoiceSlideWebOutput");
 
  //click on the answer
  SelectAnswer("ChoiceA");

  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
    
  Log.Message("Exit VerifyMultipleChoice");
}


function VerifyMultipleResponse()
{
  Log.Message("Enter VerifyMultipleResponse");
 
  //check to make sure the slide matches the known good output
  MatchSlide("MultipleResponseSlideWebOutput");
 
  //click on the answer
  SelectAnswer("ChoiceA");
  SelectAnswer("ChoiceD");

  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();   
  
  
  Log.Message("Exit VerifyMultipleResponse");  
}

function VerifyFillInTheBlank()
{
  Log.Message("Enter VerifyFillInTheBlank");

  //check to make sure the slide matches the known good output
  MatchSlide("FillInTheBlankSlideWebOutput");
 
  //click in the box
  SelectAnswer("TypeYourTextHere");

  //Figure out how to type some text....
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("choice a");
  
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();   
    
  Log.Message("Exit VerifyFillInTheBlank");    
}


function VerifyWordBank()
{
  Log.Message("Insert VerifyWordBank");

  //check to make sure the slide matches the known good output
  MatchSlide("WordBankSlideWebOutput");
 
  //drag the answer into the answer box
  DragAnswer("ChoiceADrag", "WordBankDrop");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit VerifyWordBank");
}


function VerifyMatchingDragDrop()
{
  Log.Message("Insert VerifyMatchingDragDrop");

  //check to make sure the slide matches the known good output
  MatchSlide("MatchingDragDropSlideWebOutput");
 
  //click on the answer
  DragAnswer("MatchAMDrag", "ChoiceAMDrag");
  DragAnswer("MatchBMDrag", "ChoiceBMDrag");
  DragAnswer("MatchCMDrag", "ChoiceCMDrag");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue(); 
  
  Log.Message("Exit InsertMatchingDragDrop");
}


function VerifyMatchingDropDown()
{
  Log.Message("Enter VerifyMatchingDropDown");

  Log.Message("Exit VerifyMatchingDropDown");  
}


function VerifySequenceDragAndDrop()
{
  Log.Message("Enter VerifySequenceDragAndDrop");
 
  Log.Message("Exit VerifySequenceDragAndDrop"); 
}


function VerifySequenceDropDown()
{
  Log.Message("Enter VerifySequenceDropDown");
 
  Log.Message("Exit VerifySequenceDropDown"); 
}


function VerifyNumeric()
{
  Log.Message("Insert Numeric question from Insert Questions dialog box");
 
  Log.Message("Exit InsertNumeric"); 
}

function InsertGradedHotspot()
{
  Log.Message("Insert Hotspot graded question from Insert Questions dialog box");
 
  Log.Message("Exit InsertGradedHotspot");  
}


//Survey questions
function InsertLikertScale()
{
  Log.Message("Insert Likert Scale question from Insert Questions dialog box");
   
  Log.Message("Exit InsertLikertScale");
}


function InsertPickOne()
{
  Log.Message("Insert Pick One question from Insert Questions dialog box");
   
  Log.Message("Exit InsertPickOne");
}


function InsertPickMany()
{
  Log.Message("Insert Pick Many question from Insert Questions dialog box");
  
  Log.Message("Exit InsertPickMany");
}


function InsertWhichWord()
{
  Log.Message("Insert Which Word question from Insert Questions dialog box");
  
  Log.Message("Exit InsertWhichWord");
}


function VerifyShortAnswer()
{
  Log.Message("Enter VerifyShortAnswer");

  //check to make sure the slide matches the known good output
  MatchSlide("ShortAnswerSlideWebOutput");

  //click in the box 
  SelectAnswer("TypeYourTextHere");
  
  //type some text
  
  //click submit
  ClickSubmit();
  
  Log.Message("Exit VerifyShortAnswer");
}


function InsertEssay()
{
  Log.Message("Insert Essay question from Insert Questions dialog box");
 
  Log.Message("Exit InsertEssay");
}


function InsertRankingDragDrop()
{
  Log.Message("Insert Ranking Drag and Drop question from Insert Questions dialog box");

  Log.Message("Exit InsertRankingDragDrop");
}


function InsertRankingDropDown()
{
  Log.Message("Insert Ranking Drop-down question from Insert Questions dialog box");  

  Log.Message("Exit InsertRankingDropDown");
}


function InsertHowMany()
{
  Log.Message("Insert How Many question from Insert Questions dialog box");

  Log.Message("Exit InsertHowMany");
}

//Freeform Questions
function InsertFreeformDragDrop()
{
  Log.Message("Insert Drag and Drop freeform question from Insert Questions dialog box");

  Log.Message("Exit InsertFreeFormDragDrop");
}


function InsertFreeformPickOne()
{
  Log.Message("Insert Pick One freeform question from Insert Questions dialog box");

  Log.Message("Exit InsertFreeFormPickOne");
}


function InsertFreeformPickMany()
{
  Log.Message("Insert Pick Many freeform question from Insert Questions dialog box");
 
  Log.Message("Exit InsertFreeFormPickOne");
}


function InsertTextEntry()
{
  Log.Message("Insert Text Entry question from Insert Questions dialog box");

  Log.Message("Exit InsertTextEntry");
}


function InsertFreeformHotspot()
{
  Log.Message("Insert Hotspot freeform question from Insert Questions dialog box");
 
  Log.Message("Exit InsertFreeformHotspot");  
}




