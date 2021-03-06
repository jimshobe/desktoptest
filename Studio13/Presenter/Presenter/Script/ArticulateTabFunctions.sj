function InsertCharacter()
{
  Log.Message("Enter InsertCharacter");
  
  Log.Message("insert illustrated character from ribbon");
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(391, 107);
  Aliases.POWERPNT.wndNetUIToolWindow.NetUIHWND.Click(35, 6);
  
  Log.Message("Choose character");
  Aliases.Articulate_Presenter.GalleryForm.paneCharacter.panelIllustrated.panelCharacters.tableCharacters.WaitProperty("Exists", true, 10000);
  //Aliases.Articulate_Presenter.GalleryForm.paneCharacter.panelIllustrated.panelCharacters.tableCharacters.VScroll.Pos = 0;
  //Clicks at point (256, 68) of the 'tableCharacters' object.
  Aliases.Articulate_Presenter.GalleryForm.paneCharacter.panelIllustrated.panelCharacters.tableCharacters.Click(256, 68);
  //Clicks the 'buttonInsert' button.
  Aliases.Articulate_Presenter.GalleryForm.panelStatus.buttonInsert.ClickButton();
  
  Log.Message("Exit InsertCharacter");
}


function Preview()
{
  Log.Message("Enter Preview");
  
  Log.Message("Click Preview on Ribbon");
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(807, 70);
  
  Aliases.Articulate_Presenter.MainForm.PreviewView.flashPreview.WaitProperty("Exists", "true", 10000);
  Aliases.Articulate_Presenter.MainForm.PreviewView.flashPreview.Click();
  
  Log.Message("Close preview");
  Aliases.Articulate_Presenter.MainForm.CustomRibbon.RibbonTab.SetFocus();
  Aliases.Articulate_Presenter.MainForm.CustomRibbon.RibbonTab.Click(27, 15);
  
  Log.Message("Exit Preview");
}

function PublishToWeb(title)
{
  Log.Message("Enter PublishToWeb");
  
  Log.Message("Click publish on ribbon");
  Aliases.POWERPNT.wndPPTFrameClass.panelMsodocktop.toolbarRibbon.Ribbon.NUIPane.NetUIHWND.Click(855, 73);
  Aliases.Articulate_Presenter.Publish.buttonPublish.WaitProperty("Exists", "true", 10000);
  
  Log.Message("click web target");
  Aliases.Articulate_Presenter.Publish.panelTabs.TargetTab.ClickButton();
 
  Log.Message("set title and path");
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxTitle.Select();
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxTitle.Click();
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxTitle.SetText(title);
  
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxFolder.Select();
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxFolder.Click();
  Aliases.Articulate_Presenter.Publish.panelContainer.WebTarget.textBoxFolder.SetText(Project.Variables.OutputPath);

  Log.Message("click publish and close");
  Aliases.Articulate_Presenter.Publish.buttonPublish.ClickButton();
  Aliases.Articulate_Presenter.PublishSuccess.buttonClose.WaitProperty("Exists", "true", 10000);
  Aliases.Articulate_Presenter.PublishSuccess.buttonClose.ClickButton();
  
  Log.Message("Exit PublishToWeb");
}