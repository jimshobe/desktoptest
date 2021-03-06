function MatchSlide(KnownGoodSlide)
{
  Log.Message("Enter MatchSlide");
  
  //move and resize the IE window
  Aliases.iexplore.wndIEFrame.SetFocus();
  Aliases.iexplore.wndIEFrame.Position(50,50,1350,950);
  
  //give internet explorer time to load 
  if(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).WaitProperty("Exists", "true", 50000))
  {
    var time = 0;
    do
    {
      aqUtils.Delay(100);
      time += 100;
      var slideCoordinates = Regions.Find(Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Picture(), KnownGoodSlide, 0, 0, false, false, 80);
    }
    //give the flash output/quiz time to load and keep looking for the image for 10000
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

  //find the answer on the slide and click on it
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
function SelectLikertAnswer(GivenAnswer)
{
  Log.Message("Enter SelectLikertAnswer");

  //Find the likert statement and click on the radio button show in the very right of the given image
  var theAnswer = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), GivenAnswer, 0, 0, false, false, 20);  
  if(theAnswer == null)
  {
    Log.Error("The answer can't be found.");
  }
  else
  {
    var x = theAnswer.right;
    var y = theAnswer.Top + (theAnswer.Bottom - theAnswer.Top)/2;
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(x,y);
    
    //move the mouse so it's not in the way of finding the next likert answer
    Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).HoverMouse(40,40);
   }
   
  Log.Message("Exit SelectLikertAnswer");  

}

function DragAnswer(GivenAnswer, DropLocation)
{
  Log.Message("Enter DragAnswer");

  Log.Message("Find the answer and drop location");
  var theAnswer = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), GivenAnswer, 0, 0, false, false, 50);
  var theDropLocation = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), DropLocation, 0, 0, false, false, 50);

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
    //Drag the specified answer and drop it in the specified area
    var xAnswer = theAnswer.left + (theAnswer.Right - theAnswer.Left)/2;
    var yAnswer = theAnswer.Top + (theAnswer.Bottom - theAnswer.Top)/2;
    var xDrop = theDropLocation.left + (theDropLocation.Right - theDropLocation.Left)/2;
    var yDrop = theDropLocation.Top + (theDropLocation.Bottom - theDropLocation.Top)/2;
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Drag(xAnswer, yAnswer, (xDrop-xAnswer), (yDrop-yAnswer));
   }
   
  Log.Message("Exit SelectAnswer");
}

function ClickSubmit()
{
  Log.Message("Enter ClickSubmit");

  //Find the submit button and click on it
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
    
    //Clicking on the slide removes the highlighting from the submit button on the next question slide.
    Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(150,10);

    //Move the mouse off of the flash output
    Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).HoverMouse(40,40); 
  }

  Log.Message("Exit ClickSubmit");
}


function CheckFeedback(FeedbackImage)
{
  Log.Message("Enter CheckDefaultFeedback");

  //Check to see if the specified feedback is displayed
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
  
  //Find and click on the continue button
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
    
    //Move the mouse off of the flash output
    Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).HoverMouse(40,40);
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
 
  //NOT CURRENTLY WORKING - MESSING UP WHEN THE MATCHES SHUFFLE - COME BACK AND SEE IF I CAN GET Picture.Find to work with the Region images
  /*
  //click on the answer
  var theAnswer = Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture().Find("MatchAMDrag", 0, 0, false, 0, false, 50);
  var theDropLocation = Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture().Find("ChoiceAMDrag", 0, 0, false, 0, false, 50);
  
  //var theAnswer = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), GivenAnswer, 0, 0, false, false, 50);
 // var theDropLocation = Regions.Find(Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Picture(), DropLocation, 0, 0, false, false, 50);

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
*//*
  Log.Message("Dragged match a");
  
  
  DragAnswer("MatchAMDrag", "ChoiceAMDrag");
  DragAnswer("MatchBMDrag", "ChoiceBMDrag");
  DragAnswer("MatchCMDrag", "ChoiceCMDrag");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue(); 
  */
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
  Log.Message("Enter VerifyNumeric");

  //check to make sure the slide matches the known good output
  MatchSlide("NumericSlideWebOutput");
 
  //click in the box
  SelectAnswer("NumericEntryBox");

  //Figure out how to type some text....
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("12");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit VerifyNumeric"); 
}

function VerifyGradedHotspot()
{
  Log.Message("Enter VerifyGradedHotspot");

  //check to make sure the slide matches the known good output
  MatchSlide("GradedHotspotSlideWebOutput");
 
  //click in the box
  SelectAnswer("RedTulip");

  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
   
  Log.Message("Exit VerifyGradedHotspot");  
}


//Survey questions
function VerifyLikertScale()
{
  Log.Message("Enter VerifyLikertScale");

  //check to make sure the slide matches the known good output
  MatchSlide("LikertScaleSlideWebOutput");
 
  //select answers to the likert questions
  SelectLikertAnswer("LikertStatement1");
  SelectLikertAnswer("LikertStatement2");
  SelectLikertAnswer("LikertStatement3");
  SelectLikertAnswer("LikertStatement4");

  //click submit
  ClickSubmit();
  
  Log.Message("Exit VerifyLikertScale");
}


function VerifyPickOne()
{
  Log.Message("Enter VerifyPickOne");

 //check to make sure the slide matches the known good output
  MatchSlide("PickOneSlideWebOutput");
 
  //click on the answer
  SelectAnswer("Statement1");

  //click submit
  ClickSubmit();
       
  Log.Message("Exit VerifyPickOne");
}


function VerifyPickMany()
{
  Log.Message("Enter VerifyPickMany");

 //check to make sure the slide matches the known good output
  MatchSlide("PickManySlideWebOutput");
 
  //click on the answer
  SelectAnswer("Statement1");
  SelectAnswer("Statement2");
  
  //click submit
  ClickSubmit();
      
  Log.Message("Exit VerifyPickMany");
}


function VerifyWhichWord()
{
  Log.Message("Enter VerifyWhichWord");

  //check to make sure the slide matches the known good output
  MatchSlide("WhichWordSlideWebOutput");
 
  //drag the answer into the answer box
  DragAnswer("Statement1Drag", "WordBankDrop");
  
  //click submit
  ClickSubmit();
      
  Log.Message("Exit VerifyWhichWord");
}


function VerifyShortAnswer()
{
  Log.Message("Enter VerifyShortAnswer");

  //check to make sure the slide matches the known good output
  MatchSlide("ShortAnswerSlideWebOutput");

  //click in the box 
  SelectAnswer("TypeYourTextHere");
  
  //type some text
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("This is a short answer.");
    
  //click submit
  ClickSubmit();
  
  Log.Message("Exit VerifyShortAnswer");
}


function VerifyEssay()
{
  Log.Message("Enter VerifyEssay");

  //check to make sure the slide matches the known good output
  MatchSlide("EssaySlideWebOutput");

  //click in the box 
  SelectAnswer("TypeYourTextHereEssay");
  
  //type some text
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("This is an essay. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam aliquet tellus at magna consequat rutrum. Cras ullamcorper euismod ante, sit amet semper elit rutrum a. Fusce pharetra iaculis nibh, vitae mollis arcu pretium a. Aliquam at scelerisque tortor. Suspendisse ultricies auctor ligula hendrerit tincidunt. Donec ac ante ut purus rhoncus convallis a at magna. Donec dignissim ornare ipsum ac interdum. Nam volutpat venenatis metus sit amet rhoncus. Vestibulum hendrerit urna vitae sem porta facilisis. Curabitur fringilla facilisis lacus, in convallis nulla bibendum vel. In hac habitasse platea dictumst. Aenean aliquam iaculis nibh sit amet aliquet. Proin tristique ipsum quis nisi elementum in eleifend est faucibus. Duis pharetra pellentesque adipiscing. Cras a ultrices mi. Vivamus quis tincidunt purus. Nam id arcu lorem. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam id est eros. Vivamus eu leo enim. Maecenas vitae ipsum nibh. Vivamus vel turpis eget nisl ultricies hendrerit. Nunc nisl ante, posuere at tristique cursus, eleifend ac enim. Fusce quis turpis quam. Ut non sem neque, eu varius ipsum. Fusce tristique risus non diam porta eleifend. Suspendisse suscipit facilisis tristique.");
    
  //click submit
  ClickSubmit();
  
  Log.Message("Exit VerifyEssay");
}


function VerifyRankingDragDrop()
{
  Log.Message("Enter VerifyRankingDragDrop");

  Log.Message("Exit VerifyRankingDragDrop");
}


function VerifyRankingDropDown()
{
  Log.Message("Enter VerifyRankingDropDown");  

  Log.Message("Exit VerifyRankingDropDown");
}


function VerifyHowMany()
{
  Log.Message("Enter VerifyHowMany");

  //check to make sure the slide matches the known good output
  MatchSlide("HowManySlideWebOutput");
 
  //click in the box
  SelectAnswer("NumericEntryBox");

  //Figure out how to type some text....
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("120");
  
  
  //click submit
  ClickSubmit();
    
  Log.Message("Exit VerifyHowMany");
}

//Freeform Questions
function VerifyFreeformDragDrop()
{
  Log.Message("Enter VerifyFreeformDragDrop");

  //check to make sure the slide matches the known good output
  MatchSlide("FreeformDragDropSlideWebOutput");
 
  //drag and drop the shapes onto the correct shapes
  DragAnswer("DragDropRectangle", "DragDropRoundedRectangle");
  DragAnswer("DragDropTriangle", "DragDropTrapezoid");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit VerifyFreeFormDragDrop");
}


function VerifyFreeformPickOne()
{
  Log.Message("Enter VerifyFreeformPickOne");

  //check to make sure the slide matches the known good output
  MatchSlide("FreeformPickOneSlideWebOutput");
 
  //select the shape
  SelectAnswer("PickOneRectangle");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit VerifyFreeFormPickOne");
}


function VerifyFreeformPickMany()
{
  Log.Message("Enter VerifyFreeformPickMany");

  //check to make sure the slide matches the known good output
  MatchSlide("FreeformPickManySlideWebOutput");
 
  //select the shapes
  SelectAnswer("PickManyRectangle");
  SelectAnswer("PickManySmile");
    
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
   
  Log.Message("Exit VerifyFreeFormPickOne");
}


function VerifyTextEntry()
{
  Log.Message("Enter VerifyTextEntry");

  //Click on the slide to get rid of the blinking cursor in the text entry box
  Aliases.iexplore.wndIEFrame.FrameTab.TabWindowClass.ShellDocObjectView.InternetExplorerServer.MacromediaFlashPlayerActiveX.Click(150,10);

  //move the cursor off of the slide
  Sys.Process("iexplore").Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).HoverMouse(40,40);
  
  //check to make sure the slide matches the known good output
  MatchSlide("TextEntrySlideWebOutput");
 
  //click in the box
  SelectAnswer("TypeYourTextHere");

  //Figure out how to type some text....
  Aliases.iexplore.Window("IEFrame", "QmBvt - Windows Internet Explorer", 1).Keys("answer 1");
  
  //click submit
  ClickSubmit();
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue();  
  
  Log.Message("Exit InsertTextEntry");
}


function VerifyFreeformHotspot()
{
  Log.Message("Enter VerifyFreeformHotspot");

  //check to make sure the slide matches the known good output
  MatchSlide("FreeformHotspotSlideWebOutput");
 
  //click in the box
  SelectAnswer("RedTulipBig");

  //click submit
  ClickSubmit();  
  
  //make sure the correct feedback is shown 
  CheckFeedback("DefaultCorrectFeedback");

  //click continue
  ClickContinue(); 
     
  Log.Message("Exit VerifyFreeformHotspot");  
}


function VerifyAllQuestionsResultsSlide()
{
  Log.Message("Enter VerifyAllQuestionsResultsSlide");

  //check to make sure the slide matches the known good output
  MatchSlide("AllQuestionsResultSlideWebOutput");

  Log.Message("Exit VerifyAllQuestionsResultsSlide");  

}

function CloseInternetExplorer()
{
  Log.Message("Enter CloseInternetExplorer");
  
  Aliases.iexplore.Close();
  
  Log.Message("Exit CloseInternetExplorer");
}

function ClosePublishWindow()
{
  Aliases.Quizmaker.WinFormsObject("PublishSuccess").Close();

}