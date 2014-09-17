using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerSkillsWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        private static UITestControl GetAddWorkerSkillsWindowProperties()
        {
            var addWorkerSkillsWindow = App.Container.SearchFor<WinWindow>(new { Name = "Add Worker Skills" });
            return addWorkerSkillsWindow;
        }

        private static UITestControl GetConfirmationWindowProperties()
        {
            var okWindow = App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile" });
            return okWindow;
        }

        #endregion

        #region Skill Tab Methods

        public static bool VerifyAddWorkerSkillsWindowDisplayed()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            return addWorkerSkillsWindow.Enabled;
        }

        public static void EnterDataAddWorkerSkills(DataRow data)
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();

            var postionFocus = Actions.GetWindowChild(addWorkerSkillsWindow,
                AddworkerSkillsWindowConstants.PositionFocus);
            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
            DropDownActions.SelectDropdownByText(postionFocus, data.ItemArray[12].ToString());

            var checkBox = Actions.GetWindowChild(addWorkerSkillsWindow, AddworkerSkillsWindowConstants.CheckBox);
            if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
            Actions.SetCheckBox((WinCheckBox)checkBox, data.ItemArray[14].ToString());

            MouseActions.ClickButton(addWorkerSkillsWindow, AddworkerSkillsWindowConstants.AddSelectedBtn);

            }

        public static bool ClickOnAddorUpdateBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var addBtn = Actions.GetWindowChild(workerProfileWindow, WorkerSkillsTabConstants.AddOrUpdateBtn);
                MouseActions.Click(addBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnSearchBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                var searchBtn = Actions.GetWindowChild(addWorkerSkillsWindow, AddworkerSkillsWindowConstants.SearchBtn);
                MouseActions.Click(searchBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnSaveBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                var saveBtn = Actions.GetWindowChild(addWorkerSkillsWindow, AddworkerSkillsWindowConstants.SaveBtn);
                Mouse.Click(saveBtn);
                return true;
            }
            return false;
        }

        public static void CloseWorkerSkillsWindow()
        {
            var win = (WinWindow)GetWorkerProfileWindowProperties();
            TitlebarActions.ClickClose(win);
        }

        public static bool ClickOnAddSelectedBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
            {
                var addBtn = Actions.GetWindowChild(addWorkerSkillsWindow, AddworkerSkillsWindowConstants.AddSelectedBtn);
                MouseActions.Click(addBtn);
                return true;
            }
            return false;
        }

        public static bool ClickonOkBtn()
        {
            var okWindow = GetConfirmationWindowProperties();
            if (okWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)okWindow);
                return true;
            }
            return false;

        }

        public static bool ClickonSaveBtninSkillsTab()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var saveBtn = Actions.GetWindowChild(workerProfileWindow, WorkerSkillsTabConstants.SaveBtn);
                MouseActions.Click(saveBtn);
                return true;
            }
            return false;
        }

        public static void EnterLicenseData(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var cell = TableActions.SelectCellFromTable(workerProfileWindow, "grdLicense", "Add Row", "License Type");
            cell.SetFocus();
            Mouse.DoubleClick(cell);
            SendKeys.SendWait(data.ItemArray[7].ToString());
            Actions.SendTab();
            Actions.SendText("{BACKSPACE}");
            SendKeys.SendWait(data.ItemArray[8].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[9].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[10].ToString());
            Actions.SendTab();
            SendKeys.SendWait(data.ItemArray[11].ToString());
        }

        public static void SelectDutyChkBox()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var dutyChkBox = Actions.GetWindowChild(workerProfileWindow, WorkerSkillsTabConstants.DutyChkBox);
            Actions.SetCheckBox((WinCheckBox)dutyChkBox,"TRUE");
        }

        public static void SelectVehicleChkBox()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var vChkBox = Actions.GetWindowChild(workerProfileWindow, WorkerSkillsTabConstants.VehicleChkBox);
            Actions.SetCheckBox((WinCheckBox)vChkBox, "TRUE");
        }

        #endregion

        #region Controls

        private class WorkerSkillsTabConstants
        {
            public const string AddOrUpdateBtn = "btnAddSkills";
            public const string SaveBtn = "btnSave";
            public const string DutyChkBox = "chkLightDutyWorkOnly";
            public const string VehicleChkBox = "chkVehicleAvailablity";
        }

        private class AddworkerSkillsWindowConstants
        {
            public const string PositionFocus = "cmbPositionFocus";
            public const string Title = "txtSearchText";
            public const string SkillsExperienceGrid = "grdSkillsExperience";
            public const string SearchBtn = "btnSearch";
            public const string CheckBox = "chkSelect";
            public const string AddSelectedBtn = "btnAdd";
            public const string SaveBtn = "btnAddSkillsExperience";
        }

        private class OKWindowConstants
        {
            public const string OkBtn = "-OKButton";
        }

        #endregion

    }
}