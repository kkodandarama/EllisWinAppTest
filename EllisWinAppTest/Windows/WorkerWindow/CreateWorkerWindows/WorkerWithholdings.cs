using System;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerWithholdings : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerWithholdingsWindowProperties()
        {
            var wWithholdingsWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerWithholdingsView"});//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return wWithholdingsWindow;
        }

        private static UITestControl GetW5PopUpProperties()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            var w5PopUp = wWithholdingsWindow.Container.SearchFor<WinWindow>(new {ControlId = "WorkerW5FormView"});//{ Name = "New Applicant" });
            return w5PopUp;
        }

        #endregion

        #region Withholdings Methods

        public static void ClickOnContinueBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
            {
                MouseActions.ClickButton(wWithholdingsWindow, WWithHoldingsConstants.ContinueBtn);
                Playback.Wait(3000);
            }
        }

        public static void ClickOnRejectBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
                MouseActions.ClickButton(wWithholdingsWindow, WWithHoldingsConstants.RejectBtn);
        }

        public static void ClickOnBackBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
                MouseActions.ClickButton(wWithholdingsWindow, WWithHoldingsConstants.BackBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();
            if (wWithholdingsWindow.Exists)
                MouseActions.ClickButton(wWithholdingsWindow, WWithHoldingsConstants.CancelBtn);
        }

        public static void EnterDataInWithholdings(DataRow data)
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();

            var lastName = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.LastName);
            if (!string.IsNullOrEmpty(data.ItemArray[56].ToString()))
                lastName.SetFocus();
            SendKeys.SendWait(data.ItemArray[56].ToString());

            var martialStatus = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.MaritalStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[57].ToString()))
                martialStatus.SetFocus();
            SendKeys.SendWait(data.ItemArray[57].ToString());
            SendKeys.SendWait("{TAB}");

            for (int i = 0; i < 10; i++)
            {

                if (martialStatus.Name.Equals(data.ItemArray[57].ToString()))
                {
                    i = 101;
                }
                else
                {
                    martialStatus.SetFocus();
                    SendKeys.SendWait(data.ItemArray[57].ToString());
                    SendKeys.SendWait("{TAB}");
                }

            }


            var allowances = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.Allowances);
            if (!string.IsNullOrEmpty(data.ItemArray[58].ToString()))
                allowances.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(allowances, data.ItemArray[58].ToString());

            var additional = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.AdditionalAmount);
            if (!string.IsNullOrEmpty(data.ItemArray[59].ToString()))
                additional.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(additional, data.ItemArray[59].ToString());

            var exempt = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.Exempt);
            if (!string.IsNullOrEmpty(data.ItemArray[60].ToString()))
                exempt.SetFocus();
            SendKeys.SendWait(data.ItemArray[60].ToString());

            var w5ChkBox = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.W5FormRadio);
            if (!string.IsNullOrEmpty(data.ItemArray[64].ToString()))
            {
                var checkBox = w5ChkBox as WinCheckBox;
                Factory.AssertIsFalse(checkBox == null, @"Couldn't find 'Has Worker filled out WOTC forms?' checkbox.");
                //If this checkbox is checked, we must uncheck it to get the w5PopUp.  The later if(w5PopUp.Exists) passes even if it's not visible!! (It's probably hidden, but there's no IsVisible property)
                if (checkBox.Checked)
                {
                    checkBox.SetFocus();
                    SendKeys.SendWait(" ");
                }

                Actions.SetCheckBox(checkBox, data.ItemArray[64].ToString());
            }

            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
            {
                var eligible = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Eligible);
                if (!string.IsNullOrEmpty(data.ItemArray[61].ToString()))
                    eligible.SetFocus();
                SendKeys.SendWait(data.ItemArray[61].ToString());

                var filingStatus = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Status);
                if (!string.IsNullOrEmpty(data.ItemArray[62].ToString()))
                    filingStatus.SetFocus();
                SendKeys.SendWait(data.ItemArray[62].ToString());

                var w5Spouse = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.SpouseW5);
                if (!string.IsNullOrEmpty(data.ItemArray[63].ToString()))
                    w5Spouse.SetFocus();
                SendKeys.SendWait(data.ItemArray[63].ToString());

                MouseActions.ClickButton(w5PopUp, Ww5PopupConstants.DoneBtn);
            }

            var wTocChkBox = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.WotcChk);
            if (!string.IsNullOrEmpty(data.ItemArray[65].ToString()))
                Actions.SetCheckBox((WinCheckBox)wTocChkBox, data.ItemArray[65].ToString());

            ////MouseActions.ClickButton(wWithholdingsWindow, WWithHoldingsConstants.ContinueBtn);
            ////Playback.Wait(3000);

        }

        #endregion

        #region W5 PopUp Methods

        public static void ClickOnCancelBtnW5()
        {
            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
                MouseActions.ClickButton(w5PopUp, Ww5PopupConstants.CancelBtn);
        }

        public static void ClickOnDoneBtn()
        {
            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
                MouseActions.ClickButton(w5PopUp, Ww5PopupConstants.DoneBtn);
        }

        public static void EnterDataInW5PopUp(DataRow data)
        {
            var wWithholdingsWindow = GetWorkerWithholdingsWindowProperties();

            var w5ChkBox = Actions.GetWindowChild(wWithholdingsWindow, WWithHoldingsConstants.W5FormRadio);
            var checkBox = w5ChkBox as WinCheckBox;
            Factory.AssertIsFalse(checkBox == null, @"Couldn't find checkbox in withholding window.");
            //If this checkbox is checked, we must uncheck it to get the w5PopUp.  
            if (checkBox.Checked)
            {
                checkBox.SetFocus();
                SendKeys.SendWait(" ");
            }
            if (!string.IsNullOrEmpty(data.ItemArray[64].ToString()))
                Actions.SetCheckBox((WinCheckBox)w5ChkBox, data.ItemArray[64].ToString());

            var w5PopUp = GetW5PopUpProperties();
            if (w5PopUp.Exists)
            {
                var eligible = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Eligible);
                if (!string.IsNullOrEmpty(data.ItemArray[61].ToString()))
                    eligible.SetFocus();
                SendKeys.SendWait(data.ItemArray[61].ToString());

                var filingStatus = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.Status);
                if (!string.IsNullOrEmpty(data.ItemArray[62].ToString()))
                    filingStatus.SetFocus();
                SendKeys.SendWait(data.ItemArray[62].ToString());

                var w5Spouse = Actions.GetWindowChild(w5PopUp, Ww5PopupConstants.SpouseW5);
                if (!string.IsNullOrEmpty(data.ItemArray[63].ToString()))
                    w5Spouse.SetFocus();
                SendKeys.SendWait(data.ItemArray[63].ToString());
                
            }
        }

        #endregion

        #region Controls

        private class WWithHoldingsConstants
        {
            public const string Allowances = "txtAllowances (Federal W-4 Line 5)";
            public const string AdditionalAmount = "txtAdditional Amount to WH (Federal W-4 Line 6)";
            public const string LastName = "cmbLast Name different than SSN Card? (Federal W-4 Line 4)";
            public const string MaritalStatus = "cmbMarital Status (Federal W-4 Line 3)";
            public const string Exempt = "cmbExempt (Federal W-4 Line 7)";
            public const string W5FormRadio = "chkW5Form";
            public const string WotcChk = "chkWOTCForm";
            public const string BackBtn = "_btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
        }

        private class Ww5PopupConstants
        {
            public const string Eligible = "cmbEligible (Federal W-5 Line 1)";
            public const string Status = "cmbFiling status (Federal W-5 Line 2)";
            public const string SpouseW5 = "cmbMarried and spouse has filed Form W-5 with employer (Federal W-5 Line 3)";
            public const string DoneBtn = "btnSave";
            public const string CancelBtn = "btnCancel";
        }

        #endregion
    }
}