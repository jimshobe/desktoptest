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

function LaunchEngage()
{
  try
  {
    Log.Message("Enter LaunchEngage");
  
    Log.Message("run the app");
    TestedApps.Engage.Run(1, true);
  
    //this statement is true if the activation page is displayed
    if(Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.WaitProperty("Exists", "false", 10000))
    {
      Log.Message("application is not activated");
    
      //continue if in trial period or activate if not
      if(Aliases.Engage.WizardDialog.af10_Introduction.buttonContinue.Exists)
      {
        Log.Message("In trial period so click continue");
        Aliases.Engage.WizardDialog.af10_Introduction.buttonContinue.Click();
      }
      if(Aliases.Engage.WizardDialog.af10_Introduction.buttonActivate.Exists)
      {
        Log.Message("activate application");

        Aliases.Engage.WizardDialog.af10_Introduction.buttonActivate.Click();
        Aliases.Engage.WizardDialog.af15_GetSerialNumber.textboxSerialNumber.SetText(ProjectSuite.Variables.StudioActivationKey);
      
        Aliases.Engage.WizardDialog.nextButton.WaitProperty("Exists", "true", 10000);
        Aliases.Engage.WizardDialog.nextButton.ClickButton();

        Aliases.Engage.ActivationSuccess.btnLater.WaitProperty("Exists", "true", 20000);
        Aliases.Engage.ActivationSuccess.btnLater.ClickButton();
      
        Log.Message("activated successfully");      
      }
    }

    Log.Message("Click create new project link", "");
    Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.WaitProperty("Exists", "true", 10000);
    Aliases.Engage.MainForm.ucStartPage.LinkLabelPanel1.Click();
  }
  catch(e)
  {
    Log.Error(e.description);
  }
  finally
  {
    Log.Message("Exit LaunchEngage", "");
  }
}

function InsertAccordion()
{
  try
  {
    Log.Message("Enter InsertAccordion");
  
    Log.Message("Insert accordion interaction from tab control", "");

    Aliases.Engage.NewInteractionDialog.interactionSelectionControl.imageSelectionControl.ImportGroupControl.ImportChildControl.WaitProperty("Exists", true, 10000);
    Aliases.Engage.NewInteractionDialog.interactionSelectionControl.imageSelectionControl.ImportGroupControl.ImportChildControl.Click(37, 44);
    Aliases.Engage.NewInteractionDialog.bottomPanel.okButton.ClickButton();

    Log.Message("Click introduction in left pane and insert text");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.panelFarLeft.CollapsibleControl.SidePanelHost.ItemsSidePanel.itemsListControl.panelItems.treeViewItems.Click(74, 14);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.DblClick(57, 10);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.SetText("Intro Text");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.SetText("Intro Text");

    Log.Message("Click New Rib in left panel and enter text");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.panelFarLeft.CollapsibleControl.SidePanelHost.ItemsSidePanel.itemsListControl.panelItems.treeViewItems.Click(65, 40);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.Click();
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.SetText("Rib 1");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.Click(47, 11);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.SetText("Rib Text");

    Log.Message("Click New Rib 2 and add text");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.panelFarLeft.CollapsibleControl.SidePanelHost.ItemsSidePanel.itemsListControl.panelItems.treeViewItems.Click(68, 60);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.SetText("Rib 2");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.Click();
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.SetText("Rib Text");

    Log.Message("Click New Rib 3 and add text", "");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.panelFarLeft.CollapsibleControl.SidePanelHost.ItemsSidePanel.itemsListControl.panelItems.treeViewItems.Click(75, 84);
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.textBoxTitle.SetText("Rib 3");
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.Click();
    Aliases.Engage.MainForm.panelProject.InteractionControl.ItemsPanel.panelLeft.StepPanel.editingControl.stepControl.stepTextControl.richTextBoxEx.SetText("Rib Text");
  }
  catch(e)
  {
    Log.Error(e.description)
  }
  finally
  {
    Log.Message("Exit InsertAccordion", "");
  }
}