using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerWithHoldingsWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        public static UITestControl GetWithholdingsStatePopUpProperties()
        {
            var statePopup = App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile State Withholdings Popup" });
            return statePopup;
        }

        public static UITestControl GetWorkerProfilePopUpProperties()
        {
            var workerPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile" });
            return workerPopUp;
        }

        #endregion

        #region Withholdings Tab Methods

        public static bool ClickOnSaveBtnWithholdings()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var saveBtn = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.SaveBtn);
                Mouse.Click(saveBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnAddWithholdingsBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var addBtn = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.AddWithholdingsBtn);
                Mouse.Click(addBtn);
                return true;
            }
            return false;
        }

        public static void EnterDataInWithholdings(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var lastName = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.LastName);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                lastName.SetFocus();
            SendKeys.SendWait(data.ItemArray[3].ToString());

            var martialStatus = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.MaritalStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                martialStatus.SetFocus();
            SendKeys.SendWait(data.ItemArray[4].ToString());

            var allowances = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.Allowances);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                allowances.SetFocus();
            SendKeys.SendWait("{BACKSPACE}");
            Actions.SetText(allowances, data.ItemArray[5].ToString());

            var additional = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.AdditionalAmount);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                additional.SetFocus();

            for (int i = 1; i < 10; i++)
            {
                SendKeys.SendWait("{BACKSPACE}");
            }

            Actions.SetText(additional, data.ItemArray[6].ToString());

            var exempt = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.Exempt);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                exempt.SetFocus();
            SendKeys.SendWait(data.ItemArray[7].ToString());

            var chkBox = Actions.GetWindowChild(workerProfileWindow, WorkerWithholdingsTabConstants.WotcChk);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                Actions.SetCheckBox((WinCheckBox)chkBox, data.ItemArray[8].ToString());

            MouseActions.ClickButton(workerProfileWindow, WorkerWithholdingsTabConstants.SaveBtn);
            MouseActions.ClickButton(workerProfileWindow, WorkerWithholdingsTabConstants.SaveBtn);
        }

        #endregion

        #region Controls

        public class WorkerWithholdingsTabConstants
        {
            public const string Allowances = "txtAllowances (Federal W-4 Line 5)";
            public const string AdditionalAmount = "txtAdditional Amount to WH (Federal W-4 Line 6)";
            public const string LastName = "cmbLast Name different than SSN Card? (Federal W-4 Line 4)";
            public const string MaritalStatus = "cmbMarital Status (Federal W-4 Line 3)";
            public const string Exempt = "cmbExempt (Federal W-4 Line 7)";
            public const string EffectiveStartDate = "lblExempt (Federal W-4 Line 7)";
            public const string W5FormRadio = "optW5Form";
            public const string WithholdingsGrid = "grdWithholdings";
            public const string WotcChk = "chkWOTCForm";
            public const string AddWithholdingsBtn = "btnwithholding";
            public const string SaveBtn = "btnChangesholdings";
        }

        public class WorkerStateWithholdingsPopup
        {
            public const string State = "cmbState";
            public const string WithholdingType = "cmbWithholding";
            public const string Startdate = "lblEffectiveStartDate";
            public const string Enddate = "lblEffectiveEndDate";
            public const string SubmitBtn = "btnSubmit";
            public const string CloseBtn = "btnClose";
        }

        #endregion
    }
}