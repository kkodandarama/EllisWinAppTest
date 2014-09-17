using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerPhoneWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerPhoneWindowProperties()
        {
            var phoneWindow =
                App.Container.SearchFor<WinWindow>(new { ControlId = "WorkerPhoneInformationView" });//{Name = "Phone-" + Globals.WorkerName});//WorkerPhoneInformationView
            return phoneWindow;
        }

        #endregion

        #region Phone Window Methods

        public static void EnterPhoneData(DataRow data)
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();

            var email = Actions.GetWindowChild(phoneWindow, PWorkerConstants.Email);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                Actions.SetText(email, data.ItemArray[8].ToString());

            var pPhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PPhoneNumber);
            if (!string.IsNullOrEmpty(data.ItemArray[26].ToString()))
                pPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(pPhone, data.ItemArray[26].ToString());

            var pType = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PContactType);
            if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                DropDownActions.SelectDropdownByText(pType, data.ItemArray[28].ToString());

            var pTime = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PContactTime);
            if (!string.IsNullOrEmpty(data.ItemArray[29].ToString()))
                DropDownActions.SelectDropdownByText(pTime, data.ItemArray[29].ToString());

            var pExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.PExt);
            if (!string.IsNullOrEmpty(data.ItemArray[27].ToString()))
                Actions.SetText(pExt, data.ItemArray[27].ToString());

            var sPhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SPhoneNumber);
            if (!string.IsNullOrEmpty(data.ItemArray[30].ToString()))
                sPhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(sPhone, data.ItemArray[30].ToString());

            var sType = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SContactType);
            if (!string.IsNullOrEmpty(data.ItemArray[32].ToString()))
                DropDownActions.SelectDropdownByText(sType, data.ItemArray[32].ToString());

            var sTime = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SContactTime);
            if (!string.IsNullOrEmpty(data.ItemArray[33].ToString()))
                DropDownActions.SelectDropdownByText(sTime, data.ItemArray[33].ToString());

            var sExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.SExt);
            if (!string.IsNullOrEmpty(data.ItemArray[31].ToString()))
                Actions.SetText(sExt, data.ItemArray[31].ToString());

            var ePhone = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EPhoneNumber);
            if (!string.IsNullOrEmpty(data.ItemArray[34].ToString()))
                ePhone.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            Actions.SetText(ePhone, data.ItemArray[34].ToString());

            var eName = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EContactName);
            if (!string.IsNullOrEmpty(data.ItemArray[36].ToString()))
                Actions.SetText(eName, data.ItemArray[36].ToString());

            var eRelation = Actions.GetWindowChild(phoneWindow, PWorkerConstants.ContactRelationship);
            if (!string.IsNullOrEmpty(data.ItemArray[37].ToString()))
                DropDownActions.SelectDropdownByText(eRelation, data.ItemArray[37].ToString());

            var eExt = Actions.GetWindowChild(phoneWindow, PWorkerConstants.EExt);
            if (!string.IsNullOrEmpty(data.ItemArray[35].ToString()))
                Actions.SetText(eExt, data.ItemArray[35].ToString());

            var chkBox = Actions.GetWindowChild(phoneWindow, PWorkerConstants.Application);
            if (!string.IsNullOrEmpty(data.ItemArray[75].ToString()))
                Actions.SetCheckBox((WinCheckBox)chkBox, data.ItemArray[75].ToString());

            //MouseActions.ClickButton(phoneWindow, PWorkerConstants.ContinueBtn);
            //Playback.Wait(1000);
        }

        public static void ClickOnContinueBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
            {
                MouseActions.ClickButton(phoneWindow, PWorkerConstants.ContinueBtn);
                Playback.Wait(1000);
            }
        }

        public static void ClickOnRejectBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
                MouseActions.ClickButton(phoneWindow, PWorkerConstants.RejectBtn);
        }

        public static void ClickOnBackBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
                MouseActions.ClickButton(phoneWindow, PWorkerConstants.BackBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var phoneWindow = GetWorkerPhoneWindowProperties();
            if (phoneWindow.Exists)
                MouseActions.ClickButton(phoneWindow, PWorkerConstants.CancelBtn);
        }

        #endregion

        #region Controls

        private class PWorkerConstants
        {
            public const string Email = "txtWorkerEmail";
            public const string PPhoneNumber = "mskPrimaryPhoneNumber";
            public const string PExt = "txtPrimaryExtension";
            public const string PContactType = "cmbPrimaryContactType";
            public const string PContactTime = "cmbPrimaryContactTime";
            public const string SPhoneNumber = "mskMobilePhoneNumber";
            public const string SExt = "txtMobileExtension";
            public const string SContactType = "cmbMobileContactType";
            public const string SContactTime = "cmbMobileContactTime";
            public const string EPhoneNumber = "mskEmergencyContactNumber";
            public const string EExt = "txtEmergencyExtension";
            public const string EContactName = "txtEmergencyContactName";
            public const string ContactRelationship = "cmbContactRelationship";
            public const string Application = "chkApplication";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "btnBack";
        }

        #endregion
    }
}