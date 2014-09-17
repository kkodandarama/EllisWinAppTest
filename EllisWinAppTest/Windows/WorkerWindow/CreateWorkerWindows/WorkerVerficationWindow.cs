using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerVerficationWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerVerificationWindowProperties()
        {
            var vWorkerWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerVerificationView"});//{ Name = "Verification" });
            return vWorkerWindow;
        }

        private static UITestControl GetWorkerVerificationPopUpProperties()
        {
            var vWorkerWindow =
                App.Container.SearchFor<WinWindow>(new { ControlId = "CommonDialogView"});// = "Worker Profile" });
            return vWorkerWindow;
        }

        #endregion

        #region Verification Methods

        public static void EnterVerificationData(DataRow data)
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();

            var i9Date = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.I9CompletionDt);
            if (!string.IsNullOrEmpty(data.ItemArray[39].ToString()))
                i9Date.SetFocus();
            SendKeys.SendWait(data.ItemArray[39].ToString());

            var maidenName = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.MaidenName);
            if (!string.IsNullOrEmpty(data.ItemArray[40].ToString()))
                Actions.SetText(maidenName, data.ItemArray[40].ToString());

            var dOb = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.Dob);
            if (!string.IsNullOrEmpty(data.ItemArray[41].ToString()))
                dOb.SetFocus();
            SendKeys.SendWait("^(A)");
            SendKeys.SendWait("{DEL}");
            SendKeys.SendWait(data.ItemArray[41].ToString());

            var aTitle = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.ATitle);
            if (!string.IsNullOrEmpty(data.ItemArray[48].ToString()))
                aTitle.SetFocus();
            SendKeys.SendWait(data.ItemArray[48].ToString());

            var aAuthority = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.AAuthority);
            if (!string.IsNullOrEmpty(data.ItemArray[49].ToString()))
                Actions.SetText(aAuthority, data.ItemArray[49].ToString());

            var aDocument = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.ADocument);
            if (!string.IsNullOrEmpty(data.ItemArray[50].ToString()))
                Actions.SetText(aDocument, data.ItemArray[50].ToString());
            SendKeys.SendWait("{TAB}");
            Playback.Wait(2000);
            SendKeys.SendWait(" ");

            var aExpiration = Actions.GetWindowChild(vWorkerWindow, VWorkerConstants.AExpiryDt);
            if (!string.IsNullOrEmpty(data.ItemArray[51].ToString()))
                aExpiration.SetFocus();
            SendKeys.SendWait(data.ItemArray[51].ToString());

            //MouseActions.ClickButton(vWorkerWindow, VWorkerConstants.ContinueBtn);
            //SendKeys.SendWait(" ");
            //Playback.Wait(2000);

        }

        public static void ClickOnContinueBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
            {
                MouseActions.ClickButton(vWorkerWindow, VWorkerConstants.ContinueBtn);
                //SendKeys.SendWait(" ");
                Playback.Wait(2000);
            }
        }

        public static void ClickOnRejectBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
                MouseActions.ClickButton(vWorkerWindow, VWorkerConstants.RejectBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
                MouseActions.ClickButton(vWorkerWindow, VWorkerConstants.CancelBtn);
        }

        public static void ClickOnBackBtn()
        {
            var vWorkerWindow = GetWorkerVerificationWindowProperties();
            if (vWorkerWindow.Exists)
                MouseActions.ClickButton(vWorkerWindow, VWorkerConstants.BackBtn);
        }

        public static void ClickOnOkBtnVerification()
        {
            var vWorkerWindow = GetWorkerVerificationPopUpProperties();
            if (vWorkerWindow.Exists)
            {
                MouseActions.ClickButton(vWorkerWindow, "_OKButton");
                Playback.Wait(1000);
            }
        }

        #endregion

        #region Controls

        private class VWorkerConstants
        {
            public const string I9CompletionDt = "dteI9Completion";
            public const string MaidenName = "txtMaidenName";
            public const string Dob = "dteDOB";
            public const string StatusRadioBtn = "optListValues";
            public const string DocumentsRadioBtn = "optDocumentList";
            public const string ATitle = "cmbTitle";
            public const string AAuthority = "txtAuthority";
            public const string ADocument = "txtDocument";
            public const string AExpiryDt = "dteExpDate";
            public const string CTitle = "cmbListCDocumentTitle";
            public const string CAuthority = "txtListCIssuingAuthority";
            public const string CDocument = "txtListCDocumentNumber";
            public const string CExpiryDt = "dteListCExpireDate";
            public const string Alien = "txtAlienResident";
            public const string AlienAuthorized = "txtAlienAuthorized";
            public const string I94 = "txtAdmission";
            public const string WorkUntilDt = "dteWorkUntil";
            public const string CancelBtn = "btnCancel";
            public const string ContinueBtn = "btnContinue";
            public const string BackBtn = "btnBack";
            public const string RejectBtn = "btnReject";
        }

        #endregion
    }
}