using System.Windows.Forms.VisualStyles;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class RequirementsWindow : AppContext
    {
        private static UITestControl RequirementsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Create New JobOrder"});
            //var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            return joborderWindow;
        }

        private static UITestControl AddSkillsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Add Skills"});
            return joborderWindow;
        }

        private static UITestControl AddSkillsInternalWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new {Name = "Add Skills"});
            var intwindow = joborderWindow.Container.SearchFor<WinWindow>(new {Name = "Results - Select to Add"});
            return intwindow;
        }

        private static UITestControlCollection GetAddSkillsDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = AddSkillsInternalWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetAddSkillsEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = AddSkillsInternalWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        public static void ClickOnButton(string btnName)
        {
            Factory.ClickOnButton(RequirementsWindowProperties(), btnName);
        }

        public static void EnterDatainRequirementsWindow(DataRow dataRow)
        {
            
            MouseActions.ClickButton(RequirementsWindowProperties(), "btnAddSkills");
            var windowInst = AddSkillsWindowProperties();
            if (!string.IsNullOrEmpty(dataRow.ItemArray[79].ToString()))
            {
                var ddInstance = Actions.GetWindowChild(windowInst, "cmbPositionFocus");
                DropDownActions.SelectDropdownByText(ddInstance, dataRow.ItemArray[79].ToString());
            }

            if (!string.IsNullOrEmpty(dataRow.ItemArray[80].ToString()))
            {
                var txtBoxInstance = Actions.GetWindowChild(windowInst, "txtSearchText");
                Actions.SetText(txtBoxInstance, dataRow.ItemArray[80].ToString());
                
            }

            MouseActions.ClickButton(windowInst, "btnSearch");

            
            var chkBoxControl = Actions.GetWindowChild(windowInst, "chkSelect");
            Actions.SetCheckBox((WinCheckBox) chkBoxControl, "True");

            MouseActions.ClickButton(windowInst, "btnAdd");
            MouseActions.ClickButton(windowInst, "btnAddSkillsExperience");
            
            //Click on continue button
            MouseActions.ClickButton(RequirementsWindowProperties(), "btnCancel");
            
        }
    }
}