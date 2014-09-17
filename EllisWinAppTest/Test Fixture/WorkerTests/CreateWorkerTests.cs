using System;
using System.Linq;
using System.Runtime.CompilerServices;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using EllisWinAppTest.Windows.WorkerWindow;
using EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows;
using EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.WorkerTests
{
    [CodedUITest]
    public class CreateWorkerTests : AppContext
    {
        #region Initialize Tests
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        #endregion

        #region Create Applicant Test Methods


        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CreateWorkerQualifiedTescor()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonContinueBtn();
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerConfirmApplicantElgibiltyWindow.VerifyConfirmationPopUpDisplayed(), "Confirmation PopUp Not Displayed");
                    WorkerConfirmApplicantElgibiltyWindow.ClickOnYesBtn();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void WorkerMismatchedSsn()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Mismatched SSN")
                {
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerAlreadyExistWindow.ClickonContinueBtnTelephone();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    Factory.AssertIsTrue(WorkerReveiwApplicantBehavioralSurveyResultsWindow.VerifyWorkerSurveyResultsWindowDisplayed(), "Survey Results Window Not Displayed");
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.EnterDatainTescor();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonCancelBtnTescor();
                    Playback.Wait(2000);
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void PrintNonQualifiedWorker()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[1].ToString() == "Not Qualified"))
            {
                WorkerIdentityWindow.ClickOnCreateApplicant();
                WorkerIdentityWindow.EnterApplicantData(datarow);
                WorkerIdentityWindow.ClickOnContinueBtn();
                WorkerIdentityWindow.ClickOkBtnDuplicate();
                WorkerAlreadyExistWindow.ClickOnOverideBtn();
                WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                Playback.Wait(5000);
                Factory.AssertIsTrue(WorkerReveiwApplicantBehavioralSurveyResultsWindow.VerifyWorkerNonQualifiedWindowDisplayed(), "Non Qualified Window not Displayed");
                WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonPrintBtn();
                Playback.Wait(1000);
                WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClosePrintWindow();
                WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonCloseBtn();
            }
            Cleanup();

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CloseNonQualifiedWorker()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Not Qualified")
                {
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerIdentityWindow.ClickOkBtnDuplicate();
                    WorkerAlreadyExistWindow.ClickOnOverideBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    Factory.AssertIsTrue(WorkerReveiwApplicantBehavioralSurveyResultsWindow.VerifyWorkerNonQualifiedWindowDisplayed(), "Non Qualified Window not Displayed");
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickonCloseBtn();
                }
            }
            Cleanup();

        }

        #endregion

        #region Update Applicant Test Methods

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void UpdateWorker()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    var updateBtn = WorkerAlreadyExistWindow.VerifyUpdateBtnEnabled();
                    if (updateBtn)
                    {
                        WorkerAlreadyExistWindow.ClickOnUpdateProfileBtn();
                        Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Summary Tab not Displayed");
                        WorkerSummaryWindow.ClickOnCloseBtn();
                    }
                    else
                    {
                        Factory.AssertIsTrue(updateBtn, "No workers Displayed to Update");
                        WorkerAlreadyExistWindow.ClickOnBackBtn();
                        WorkerIdentityWindow.ClickOnCancelBtn();
                        WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    }
                }
            }
            Cleanup();
        }

        #endregion

        #region Override Applicant Test Methods

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void OverrideWorker()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    var overRideBtn = WorkerAlreadyExistWindow.VerifyOverRideBtnEnabled();
                    if (overRideBtn)
                    {
                        WorkerAlreadyExistWindow.ClickOnOverideBtn();
                        Factory.AssertIsTrue(WorkerCompleteBehavioralSurveryWindow.VerifyCompleteBehavioralSurveyWindowDisplayed(), "Survey Window Not displayed");
                    }
                    else
                    {
                        Factory.AssertIsTrue(overRideBtn, "No Workers Displayed To Override");
                        WorkerAlreadyExistWindow.ClickOnBackBtn();
                        WorkerIdentityWindow.ClickOnCancelBtn();
                        WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    }
                }
            }
            Cleanup();
        }

        #endregion

        #region Cancel Applicant Test Methods

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CancelWorkerAddressData()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered Address data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CancelWorkerEmploymentEligibilityData()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered Verification data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void ClickBackBtnWorkerEmploymentEligibilityData()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnBackBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered Verification data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    WorkerPhoneWindow.ClickOnCancelBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();

                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void ClickCancelBtnAlertWindow()
        {
            Initialize();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
            foreach (var datarow in datarows)
            {
                if (datarow.ItemArray[1].ToString() == "Qualified")
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered Identity data is not cancelled");
                    WorkerIdentityWindow.ClickOnCancelBtnPopUp();
                    WorkerIdentityWindow.ClickOnCancelBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CancelWorkerW5Screen()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInW5PopUp(datarow);
                    WorkerWithholdings.ClickOnCancelBtnW5();
                    WorkerWithholdings.ClickOnCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered WithHoldings data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CancelWorkerWithHoldingsScreen()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered WithHoldings data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void CancelWorkerJobSkills()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonCancelBtn();
                    Factory.AssertIsTrue(WorkerIdentityWindow.VerifyAlertPopUpDisplayed(), "Alert PopUp Not Displayed. Entered Job Skills data is not cancelled");
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        //[TestMethod]
        //[TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        //public void CancelWorkerConfirmationScreen()
        //{
        //    Initialize();

        //    var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
        //    foreach (var datarow in datarows)
        //    {
        //        if (datarow.ItemArray[1].ToString() == "Qualified")
        //        {
        //            Playback.Wait(1000);
        //            WorkerIdentityWindow.ClickOnCreateApplicant();
        //            WorkerIdentityWindow.EnterApplicantData(datarow);
        //            WorkerIdentityWindow.ClickOnContinueBtn();
        //            WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
        //            WorkerAlreadyExistWindow.ClickOnContinueBtn();
        //            WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
        //            Playback.Wait(5000);
        //            WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
        //            WorkerAddressWindow.EnterAddressData(datarow);
        //            WorkerAddressWindow.ClickOnContinueBtn();
        //            WorkerGeoCodeWindow.ClickOnOkBtn();
        //            WorkerVertexGeoCodeWindow.ClickOnOkBtn();
        //            Playback.Wait(2000);
        //            WorkerPhoneWindow.EnterPhoneData(datarow);
        //            WorkerPhoneWindow.ClickOnContinueBtn();
        //            WorkerVerficationWindow.EnterVerificationData(datarow);
        //            WorkerVerficationWindow.ClickOnContinueBtn();
        //            WorkerVerficationWindow.ClickOnOkBtnVerification();
        //            WorkerWithholdings.EnterDataInWithholdings(datarow);
        //            WorkerWithholdings.ClickOnContinueBtn();
        //            WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
        //            WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
        //            WorkerJobSkillsWindow.EnterLicenseData(datarow);
        //            WorkerJobSkillsWindow.ClickonContinueBtn();
        //            Playback.Wait(2000);
        //            WorkerConfirmApplicantElgibiltyWindow.ClickOnNoBtn();
        //            Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(),"Reject PopUp Not Displayed");
        //            WorkerAddressWindow.EnterDataInRejectPopUp();
        //            WorkerAddressWindow.ClickOnDoneBtnReject();
        //        }
        //    }

        //    Cleanup();
        //}


        #endregion

        #region Reject Applicant Test Methods

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RejectWorker()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString() == "Qualified"))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnRejectBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(), "Reject PopUp Not Displayed");
                    WorkerAddressWindow.EnterDataInRejectPopUp();
                    WorkerAddressWindow.ClickOnDoneBtnReject();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RejectWorkerVerification()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnRejectBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(), "Reject PopUp Not Displayed");
                    WorkerAddressWindow.EnterDataInRejectPopUp();
                    WorkerAddressWindow.ClickOnDoneBtnReject();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RejectWorkerW5Screen()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnRejectBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(), "Reject PopUp Not Displayed");
                    WorkerAddressWindow.EnterDataInRejectPopUp();
                    WorkerAddressWindow.ClickOnDoneBtnReject();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RejectWorkerJobSkills()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(x => x.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonRejectBtn();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    WorkerIdentityWindow.ClickOnOkBtnPopUp();
                    Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(), "Reject PopUp Not Displayed");
                    WorkerAddressWindow.EnterDataInRejectPopUp();
                    WorkerAddressWindow.ClickOnDoneBtnReject();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RejectWorkerConfirmationScreen()
        {
            try
            {
                Initialize();

                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateWorker);
                foreach (var datarow in datarows.Where(v => v.ItemArray[1].ToString().Equals("Qualified")))
                {
                    Playback.Wait(1000);
                    WorkerIdentityWindow.ClickOnCreateApplicant();
                    WorkerIdentityWindow.EnterApplicantData(datarow);
                    WorkerIdentityWindow.ClickOnContinueBtn();
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnNextBtnQualified();
                    WorkerAlreadyExistWindow.ClickOnContinueBtn();
                    WorkerCompleteBehavioralSurveryWindow.ClickOnGetResultsBtn();
                    Playback.Wait(5000);
                    WorkerReveiwApplicantBehavioralSurveyResultsWindow.ClickOnContinueBtn();
                    WorkerAddressWindow.EnterAddressData(datarow);
                    WorkerAddressWindow.ClickOnContinueBtn();
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Playback.Wait(2000);
                    WorkerPhoneWindow.EnterPhoneData(datarow);
                    WorkerPhoneWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.EnterVerificationData(datarow);
                    WorkerVerficationWindow.ClickOnContinueBtn();
                    WorkerVerficationWindow.ClickOnOkBtnVerification();
                    WorkerWithholdings.EnterDataInWithholdings(datarow);
                    WorkerWithholdings.ClickOnContinueBtn();
                    WorkerJobSkillsWindow.ClickonAddOrUpdateBtn();
                    WorkerJobSkillsWindow.EnterDataInAddSkills(datarow);
                    WorkerJobSkillsWindow.EnterLicenseData(datarow);
                    WorkerJobSkillsWindow.ClickonContinueBtn();
                    Playback.Wait(2000);
                    WorkerConfirmApplicantElgibiltyWindow.ClickOnNoBtn();
                    Factory.AssertIsTrue(WorkerAddressWindow.VerifyRejectPopUpDisplayed(), "Reject PopUp Not Displayed");
                    WorkerAddressWindow.EnterDataInRejectPopUp();
                    WorkerAddressWindow.ClickOnDoneBtnReject();
                }
            }
            finally
            {
                Cleanup();
            }

        }

        #endregion

        #region Cleanup Tests

        public
            void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

        #endregion

    }
}
