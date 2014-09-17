using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerAddressWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerAddressWindowProperties()
        {
            var wAddresWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerContactInformationView"});//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });//WorkerContactInformationView
            return wAddresWindow;
        }

        private static UITestControl GetRejectPopUpProperties()
        {
            var rejectPopUp =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerRejectionView"});//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });//WorkerRejectionView
            return rejectPopUp;
        }

        private static UITestControl GetAlertSkillsProperties()
        {
            var alertPopUp =
                App.Container.SearchFor<WinWindow>(new { Name = "Alert" });
            return alertPopUp;
        }

        #endregion

        #region Address Methods

        public static void EnterAddressData(DataRow data)
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();

            var suffix = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Suffix);
            if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
                DropDownActions.SelectDropdownByText(suffix, data.ItemArray[11].ToString());

            var gender = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Gender);
            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                DropDownActions.SelectDropdownByText(gender, data.ItemArray[12].ToString());

            var race = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.Race);
            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                DropDownActions.SelectDropdownByText(race, data.ItemArray[13].ToString());

            var jobCategory = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.JobCategory);
            if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
                DropDownActions.SelectDropdownByText(jobCategory, data.ItemArray[14].ToString());

            var rAddress = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RStreetAddress);
            if (!string.IsNullOrEmpty(data.ItemArray[15].ToString()))
                Actions.SetText(rAddress, data.ItemArray[15].ToString());

            var rSuiteNo = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RApartment);
            if (!string.IsNullOrEmpty(data.ItemArray[16].ToString()))
                Actions.SetText(rSuiteNo, data.ItemArray[16].ToString());

            var rState = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RState);
            if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
                DropDownActions.SelectDropdownByText(rState, data.ItemArray[17].ToString());

            var rZip = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RZip);
            if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
                DropDownActions.SelectDropdownByText(rZip, data.ItemArray[18].ToString());

            var rCity = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.RCity);
            if (!string.IsNullOrEmpty(data.ItemArray[19].ToString()))
                DropDownActions.SelectDropdownByText(rCity, data.ItemArray[19].ToString());

            var mAddress = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MStreetAddress);
            if (!string.IsNullOrEmpty(data.ItemArray[20].ToString()))
                Actions.SetText(mAddress, data.ItemArray[20].ToString());

            var mSuiteNo = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MApartment);
            if (!string.IsNullOrEmpty(data.ItemArray[21].ToString()))
                Actions.SetText(mSuiteNo, data.ItemArray[21].ToString());

            var mState = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MState);
            if (!string.IsNullOrEmpty(data.ItemArray[22].ToString()))
                DropDownActions.SelectDropdownByText(mState, data.ItemArray[22].ToString());

            var mZip = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MZip);
            if (!string.IsNullOrEmpty(data.ItemArray[23].ToString()))
                DropDownActions.SelectDropdownByText(mZip, data.ItemArray[23].ToString());

            var mCity = Actions.GetWindowChild(wAddresWindow, AWorkerConstants.MCity);
            if (!string.IsNullOrEmpty(data.ItemArray[24].ToString()))
                DropDownActions.SelectDropdownByText(mCity, data.ItemArray[24].ToString());

            //MouseActions.ClickButton(wAddresWindow, AWorkerConstants.ContinueBtn);
            //Playback.Wait(1000);
        }

        public static void ClickOnContinueBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
            {
                MouseActions.ClickButton(wAddresWindow, AWorkerConstants.ContinueBtn);
                Playback.Wait(1000);
            }
        }

        public static void ClickOnBackBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
                MouseActions.ClickButton(wAddresWindow, AWorkerConstants.BackBtn);
        }

        public static void ClickOnRejectBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
                MouseActions.ClickButton(wAddresWindow, AWorkerConstants.RejectBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var wAddresWindow = GetWorkerAddressWindowProperties();
            if (wAddresWindow.Exists)
                MouseActions.ClickButton(wAddresWindow, AWorkerConstants.CancelBtn);
        }

        #endregion

        #region Reject PopUp Methods

        public static void ClickOnDoneBtnReject()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
                MouseActions.ClickButton(rejectPopUp, "btnDone");
        }

        public static void ClickOnBackBtnReject()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
                MouseActions.ClickButton(rejectPopUp, "btnBack");
        }

        public static void EnterDataInRejectPopUp()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            if (rejectPopUp.Exists)
            {
                var rejectCmb = Actions.GetWindowChild(rejectPopUp, "ultraRejectCombo");
                DropDownActions.SelectDropdownByText(rejectCmb, "Others");
            }
        }

        public static bool VerifyRejectPopUpDisplayed()
        {
            var rejectPopUp = GetRejectPopUpProperties();
            return rejectPopUp.Exists;
        }

        public static bool VerifyJobSkillsAlertDisplayed()
        {
            var alertSkills = GetAlertSkillsProperties();
            return alertSkills.Exists;
        }

        #endregion

        #region Controls

        private class AWorkerConstants
        {
            public const string Suffix = "cmbSuffix";
            public const string Gender = "cmbGender";
            public const string Race = "cmbRace";
            public const string JobCategory = "cmbJobCategory";
            public const string RStreetAddress = "txtResidenceStreet";
            public const string RApartment = "txtResidenceApt";
            public const string RState = "cmbResidenceState";
            public const string RZip = "cmbResidenceZipcode";
            public const string RCity = "cmbResidenceCity";
            public const string MStreetAddress = "txtMailingStreet";
            public const string MApartment = "txtMailingApt";
            public const string MState = "cmbMailingState";
            public const string MZip = "cmbMailingZipCode";
            public const string MCity = "cmbMailingCity";
            public const string BackBtn = "_btnBack";
            public const string CancelBtn = "btnCancel";
            public const string RejectBtn = "btnReject";
            public const string ContinueBtn = "btnContinue";
            public const string AddressChkBox = "chkAddress";
            public const string DeliveryRadioBtn = "OptDelivery";
        }

        #endregion
    }
}