using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Web;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerIdentityWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetCreateApplicantWindowProperties()
        {
            var applicantWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Identity" });
            return applicantWindow;
        }

        private static UITestControl DuplicateEmailAddressPopUpProperties()
        {
            var duplicateEmailAddressPopUp = App.Container.SearchFor<WinWindow>(new { Name = "Duplicate Email" });
            return duplicateEmailAddressPopUp;
        }

        private static UITestControl AlertPopUpProperties()
        {
            var popup = App.Container.SearchFor<WinWindow>(new { Name = "Alert" });
            return popup;
        }

        #endregion

        #region Identity Tab Methods

        public static void ClickOnCreateApplicant()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.File });
            var worker = file.Container.SearchFor<WinMenuItem>(new { Name = "Workers" });
            var applicant = worker.Container.SearchFor<WinMenuItem>(new { Name = "Create Applicant" });


            MouseActions.Click(file);
            MouseActions.Click(worker);
            MouseActions.Click(applicant);
        }

        public static void ClickOnContinueBtn()
        {
            var applicantWindow = GetCreateApplicantWindowProperties();
            if (applicantWindow.Exists)
            {
                MouseActions.ClickButton(applicantWindow, IWorkerConstants.ContinueBtn);
                Playback.Wait(1000);
            }
        }

        public static void ClickOnCancelBtn()
        {
            var applicantWindow = GetCreateApplicantWindowProperties();
            if (applicantWindow.Exists)
                MouseActions.ClickButton(applicantWindow, IWorkerConstants.CancelBtn);
        }

        public static void ClickOkBtnDuplicate()
        {
            var duplicateEmailAddressPopUp = DuplicateEmailAddressPopUpProperties();
            if (duplicateEmailAddressPopUp.Exists)
                MouseActions.ClickButton(duplicateEmailAddressPopUp, IWorkerConstants.OkBtn);
        }

        public static bool VerifyDuplicateEmailAddressWindowDisplayed()
        {
            var duplicateEmailAddressPopUp = DuplicateEmailAddressPopUpProperties();

            if (duplicateEmailAddressPopUp.Exists)
            {
                return true;
            }
            return false;
        }

        public static void EnterApplicantData(DataRow data)
        {
            var applicantWindow = GetCreateApplicantWindowProperties();

            var firstName = Actions.GetWindowChild(applicantWindow, IWorkerConstants.FirstName);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                Actions.SetText(firstName, data.ItemArray[3].ToString());

            var middleInitial = Actions.GetWindowChild(applicantWindow, IWorkerConstants.MiddleInitial);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                Actions.SetText(middleInitial, data.ItemArray[4].ToString());

            var lastName = Actions.GetWindowChild(applicantWindow, IWorkerConstants.LastName);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                Actions.SetText(lastName, data.ItemArray[5].ToString());

            var ssn = Actions.GetWindowChild(applicantWindow, IWorkerConstants.SSN);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                Actions.SetText(ssn, data.ItemArray[6].ToString());

            var phone = Actions.GetWindowChild(applicantWindow, IWorkerConstants.PrimaryPhone);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                Actions.SetText(phone, data.ItemArray[7].ToString());

            var contactType = Actions.GetWindowChild(applicantWindow, IWorkerConstants.ContactType);
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                if(contactType.Enabled)
                DropDownActions.SelectDropdownByText(contactType, data.ItemArray[9].ToString());

            var email = Actions.GetWindowChild(applicantWindow, IWorkerConstants.Email);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                Actions.SetText(email, data.ItemArray[8].ToString());

            var laborReady = Actions.GetWindowChild(applicantWindow, IWorkerConstants.LaborReady);
            laborReady.SetFocus();
            if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                SendKeys.SendWait(data.ItemArray[10].ToString());

            //MouseActions.ClickButton(applicantWindow, IWorkerConstants.ContinueBtn);
        }

        #endregion

        #region Alert PopUp Methods

        public static void ClickOnOkBtnPopUp()
        {
            var popup = AlertPopUpProperties();
            if (popup.Exists)
                MouseActions.ClickButton(popup, "_OKButton");
        }

        public static void ClickOnCancelBtnPopUp()
        {
            var popup = AlertPopUpProperties();
            if (popup.Exists)
                MouseActions.ClickButton(popup, "_CancelButton");
        }

        public static bool VerifyAlertPopUpDisplayed()
        {
            var popup = AlertPopUpProperties();
            return popup.Exists;
        }

        #endregion

        #region Controls

        private class IWorkerConstants
        {
            public const string FirstName = "txtFirstName";
            public const string MiddleInitial = "txtMiddleInitial";
            public const string LastName = "txtLastName";
            public const string SSN = "mskSSN";
            public const string PrimaryPhone = "mskPhone";
            public const string Email = "txtJobSiteEmail";
            public const string ContactType = "cmbMobileContactType";
            public const string LaborReady = "cmbLaborReady";
            public const string CancelBtn = "btnCancel";
            public const string ContinueBtn = "btnContinue";
            public const string OkBtn = "btnOk";
        }

        #endregion
    }
}

