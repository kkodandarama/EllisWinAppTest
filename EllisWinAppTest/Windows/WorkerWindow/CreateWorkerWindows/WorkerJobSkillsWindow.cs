using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerJobSkillsWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerSkillsWindowProperties()
        {
            var jWorkerWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerJobSkillsView"});//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return jWorkerWindow;
        }

        private static UITestControl GetAddWorkerSkillsWindowProperties()
        {
            var addWorkerSkillsWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerSkillsAndExperienceView"});//{ Name = "Add Worker Skills" });
            return addWorkerSkillsWindow;
        }

        #endregion

        #region Job Skills Methods

        public static void ClickonAddOrUpdateBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();

            MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.AddOrUpdateBtn);
            Playback.Wait(1000);
        }

        public static void ClickonBackBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.BackBtn);
        }

        public static void ClickonCancelBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.CancelBtn);
        }

        public static void ClickonRejectBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.RejectBtn);
        }

        public static void ClickonContinueBtn()
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();
            if (jWorkerWindow.Exists)
            {
                MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.ContinueBtn);
                Playback.Wait(1000);
            }
        }

        //public static void EnterLicenseData(DataRow data)
        //{
        //    var jWorkerWindow = GetWorkerSkillsWindowProperties();

        //    var cell = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "Add Row", "License Type");
        //    cell.SetFocus();
        //    Mouse.DoubleClick(cell);
        //    SendKeys.SendWait(data.ItemArray[68].ToString());
        //    Actions.SendTab();
        //    Playback.Wait(1000);
        //    Actions.SendText("{BACKSPACE}");
        //    SendKeys.SendWait(data.ItemArray[69].ToString());
        //    Actions.SendTab();
        //    Playback.Wait(1000);
        //    SendKeys.SendWait(data.ItemArray[70].ToString());
        //    Actions.SendTab();
        //    Playback.Wait(1000);
        //    SendKeys.SendWait(data.ItemArray[71].ToString());
        //    Actions.SendTab();
        //    Playback.Wait(1000);
        //    SendKeys.SendWait(data.ItemArray[72].ToString());
        //    Playback.Wait(2000);

        //    var dutyChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.DutyChkBox);
        //    Actions.SetCheckBox((WinCheckBox)dutyChkBox, data.ItemArray[73].ToString());

        //    var vehicleChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.VehicleChkBox);
        //    Actions.SetCheckBox((WinCheckBox)vehicleChkBox, data.ItemArray[74].ToString());

        //    MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.ContinueBtn);
        //    Playback.Wait(1000);
        //}

        public static void EnterLicenseData(DataRow data)
        {
            var jWorkerWindow = GetWorkerSkillsWindowProperties();

            var lType = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "WorkerLicensesDomain row 1", "License Type");
            if (!string.IsNullOrEmpty(data.ItemArray[68].ToString()))
                lType.SetFocus();
            Mouse.Click(lType);
            SendKeys.SendWait(data.ItemArray[68].ToString());
            Playback.Wait(1000);

            var eDate = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "WorkerLicensesDomain row 1", "Expiration Date");
            if (!string.IsNullOrEmpty(data.ItemArray[69].ToString()))
                Actions.SetText(eDate, data.ItemArray[69].ToString());
            Playback.Wait(1000);

            var issue = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "WorkerLicensesDomain row 1", "Issued By");
            if (!string.IsNullOrEmpty(data.ItemArray[70].ToString()))
                issue.SetFocus();
            Mouse.Click(issue);
            SendKeys.SendWait(data.ItemArray[70].ToString());
            Playback.Wait(1000);

            var state = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "WorkerLicensesDomain row 1", "Federal/State");
            if (!string.IsNullOrEmpty(data.ItemArray[71].ToString()))
                state.SetFocus();
            Mouse.Click(state);
            SendKeys.SendWait(data.ItemArray[71].ToString());
            Playback.Wait(1000);

            var moreInfo = TableActions.SelectCellFromTable(jWorkerWindow, "grdLicense", "WorkerLicensesDomain row 1", "More Info");
            if (!string.IsNullOrEmpty(data.ItemArray[72].ToString()))
                Actions.SetText(moreInfo, data.ItemArray[72].ToString());
            Playback.Wait(2000);

            var dutyChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.DutyChkBox);
            Actions.SetCheckBox((WinCheckBox)dutyChkBox, data.ItemArray[73].ToString());

            var vehicleChkBox = Actions.GetWindowChild(jWorkerWindow, WJobSkillsWindowConstants.VehicleChkBox);
            Actions.SetCheckBox((WinCheckBox)vehicleChkBox, data.ItemArray[74].ToString());

            //MouseActions.ClickButton(jWorkerWindow, WJobSkillsWindowConstants.ContinueBtn);
            //Playback.Wait(1000);
        }

        #endregion

        #region Add Worker Skills Methods

        public static void ClickonSearchBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.SearchBtn);
        }

        public static void ClickonAddSelectedBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.AddSelectedBtn);
        }

        public static void ClickonSaveBtn()
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();
            if (addWorkerSkillsWindow.Exists)
                MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.SaveBtn);
        }

        public static void EnterDataInAddSkills(DataRow data)
        {
            var addWorkerSkillsWindow = GetAddWorkerSkillsWindowProperties();

            var postionFocus = Actions.GetWindowChild(addWorkerSkillsWindow,
            WorkerAddSkillsWindowConstants.PositionFocus);
            var cmbBox = (WinComboBox)postionFocus;
            if (!string.IsNullOrEmpty(data.ItemArray[66].ToString()))
                DropDownActions.SelectDropdownByText(cmbBox, data.ItemArray[66].ToString());

            var checkBox = Actions.GetWindowChild(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.CheckBox);
            if (!string.IsNullOrEmpty(data.ItemArray[67].ToString()))
                Actions.SetCheckBox((WinCheckBox)checkBox, data.ItemArray[67].ToString());

            MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.AddSelectedBtn);

            MouseActions.ClickButton(addWorkerSkillsWindow, WorkerAddSkillsWindowConstants.SaveBtn);
            Playback.Wait(1000);

        }

        #endregion

        #region Controls

        private class WJobSkillsWindowConstants
        {
            public const string AddOrUpdateBtn = "btnAddSkills";
            public const string BackBtn = "btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string DutyChkBox = "chkLightDutyWorkOnly";
            public const string VehicleChkBox = "chkVehicleAvailablity";
            public const string LicenseGrid = "grdLicenses";
            public const string CertificationsGrid = "grdCertifications";
            public const string BackgroundCheckGrid = "grdBackgroundCheck";
            public const string SkillsGrid = "ultraGrid1";

        }

        private class WorkerAddSkillsWindowConstants
        {
            public const string PositionFocus = "cmbPositionFocus";
            public const string Title = "txtSearchText";
            public const string SkillsExperienceGrid = "grdSkillsExperience";
            public const string SearchBtn = "btnSearch";
            public const string CheckBox = "chkSelect";
            public const string AddSelectedBtn = "btnAdd";
            public const string SaveBtn = "btnAddSkillsExperience";
        }

        #endregion

    }
}
