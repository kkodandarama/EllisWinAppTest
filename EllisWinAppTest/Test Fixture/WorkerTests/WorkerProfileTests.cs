using System;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Elements;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows;
using EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.WorkerTests
{
    [CodedUITest]
    public class WorkerProfileTests : AppContext
    {

        #region Initialize Tests
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        #endregion

        #region Summary Tab Tests

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerSummaryTab()
        {
            try
            {
                Initialize();

                LandingPage.SelectFromToolbar("Workers");
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
                foreach (var datarow in datarows)
                {
                    var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                    if (worker)
                    {
                        Playback.Wait(2000);
                        Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                            "Worker Summary Tab not Displayed");
                        WorkerSummaryWindow.ClickOnCloseBtn();
                    }
                    Factory.AssertIsTrue(worker, "Requested Worker not found");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void ClickOnCloseBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Summary Tab not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void ClickOnChangeStatusBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Summary Tab not Displayed");
                    WorkerSummaryWindow.ClickOnChangeStatusBtn();
                    Factory.AssertIsTrue(WorkerChangeStatusWindow.VerifyChangeStatusWindowDisplayed(),
                        "Change Status Window Not Displayed");
                    WorkerChangeStatusWindow.ClickOnCancelBtnStatusWindow();
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSummaryWindow.VerifyAlertPopUpDisplayed(), "Alert Pop Up Not Displayed");
                    WorkerSummaryWindow.CloseAlertPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void ClickOnPrintReportBtnWorkerProfile()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Summary Tab not Displayed");
                    WorkerSummaryWindow.ClickOnPrintReportBtn();
                    Playback.Wait(2000);
                    WorkerSummaryWindow.ClosePrintWindow();
                    Playback.Wait(2000);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyDropDownValuesChangeStatusWindow()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Summary Tab Not Displayed");
                    WorkerSummaryWindow.ClickOnChangeStatusBtn();
                    Factory.AssertIsTrue(WorkerChangeStatusWindow.VerifyChangeStatusWindowDisplayed(),
                        "Change Status Window Not Displayed");
                    WorkerChangeStatusWindow.ClickOnStatusDropDown();
                    Playback.Wait(2000);
                    WorkerChangeStatusWindow.ClickOnCancelBtnStatusWindow();
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSummaryWindow.VerifyAlertPopUpDisplayed(), "Alert Pop Up Not Displayed");
                    WorkerSummaryWindow.CloseAlertPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        #endregion

        #region Profile Details Tab Tests

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerProfileFromActiveWorkers()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerIdentity);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[8].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Window not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerProfileTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerIdentity);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[8].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Tab not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerIdentityTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerIdentity);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[8].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerProfileDetailsWindow.EnterdataInIdentity(datarow);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();

        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerContactsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            Factory.AssertIsTrue(WorkerSummaryWindow.VerifyWorkersDisplayed(), "Worker Landing Page not Displayed");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerContacts);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[20].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Contacts");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Contacts Tab not Displayed");
                    WorkerProfileDetailsWindow.EnterDataInContactTabs(datarow);
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerAddressTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAddress);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[14].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Addresses");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Addresses Tab not Displayed");
                    WorkerProfileDetailsWindow.EnterDataInAddressTab(datarow);
                    Factory.AssertIsTrue(WorkerGeoCodeWindow.VerifyGeoCodeWindowDisplayed(), "Geo Code Window Not Displayed");
                    WorkerGeoCodeWindow.ClickOnOkBtn();
                    Factory.AssertIsTrue(WorkerVertexGeoCodeWindow.VerifyWorkerVertexGeoCodeWindowDisplayed(),
                        " Worker Vertex Geo Code Window Not Displayed");
                    WorkerVertexGeoCodeWindow.ClickOnOkBtn();
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }




        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerTempToHireTab()
        {
            try
            {
                Initialize();
                LandingPage.SelectFromToolbar("Workers");
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerTemp);
                var datarow = datarows.FirstOrDefault(x => !(x.ItemArray[7].ToString().Equals(String.Empty)));
                Factory.AssertIsFalse(datarow == null, "Couldn't find a worker in the data.");
                Globals.WorkerName = datarow.ItemArray[7].ToString();
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(Globals.WorkerName);
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Temp-to-Hire");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Temp-to-Hire Tab not Displayed");
                    WorkerProfileDetailsWindow.ClickonSearchBtnTemp();
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyCustomerSearchPopUpDisplayed(),
                        "Customer Search PopUp not Displayed");
                    WorkerProfileDetailsWindow.EnterDatainCustomerSearch(datarow);
                    Playback.Wait(1000);
                    WorkerProfileDetailsWindow.EnterdataInTemp(datarow);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(), "Confirmation Window Not Displayed");
                    //WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");

            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerAvailabilityTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerTemp);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[7].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Availability");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Availability Tab not Displayed");
                    WorkerProfileDetailsWindow.EnterdataInAvailability(datarow);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerVerificationTab()
        {
            try
            {
                Initialize();
                LandingPage.SelectFromToolbar("Workers");
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerVerification);
                foreach (var datarow in datarows.Where(x => !x.ItemArray[19].ToString().Equals(string.Empty)))
                {
                    var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[19].ToString());
                    if (worker)
                    {
                        Playback.Wait(2000);
                        WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                        WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                            "I-9");
                        Playback.Wait(1000);
                        //Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Profile Verification Tab not Displayed");
                        WorkerProfileDetailsWindow.EnterDatainVerification(datarow);
                        WorkerProfileDetailsWindow.ClickOnSaveBtnVerification();
                        Playback.Wait(1000);
                        Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                            "Confirmation Window Not Displayed");
                        WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                        Playback.Wait(1000);
                        WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                        Playback.Wait(1000);
                        WorkerSummaryWindow.ClickOnCloseBtn();
                    }
                    Factory.AssertIsTrue(worker, "Requested Worker not found");
                }
            }
            finally
            {
                Cleanup();
            }

        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerPaymentMethodTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerVerification);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[19].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Payment Method");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Payment Method Tab not Displayed");
                    SelectRadioButton.Selection(datarow.ItemArray[18].ToString());
                    WorkerProfileDetailsWindow.ClickOnEditBankDetailsBtn();
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyBankPopUpDisplayed(),
                        "Bank PopUp Not Displayed");
                    WorkerProfileDetailsWindow.EnterDataInBankPopUp(datarow);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyBankConfirmationPopUpDisplayed(),
                        "Bank Confirmation Pop Up Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBankConfirmation();
                    WorkerProfileDetailsWindow.ClickOnSaveBtnPayment();
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerStatusTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerVerification);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[19].ToString());
                if (worker)
                {
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails);
                    WorkerSummaryWindow.SelectInnerTab(WorkerSummaryWindow.WorkerProfileTabConstants.ProfileDetails,
                        "Status");
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Profile Status Tab not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        #endregion

        #region Withholdings Tab Tests

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerWithholdingsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerWithholdings);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[14].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Withholdings);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Withholdings Tab not Displayed");
                    WorkerWithHoldingsWindow.EnterDataInWithholdings(datarow);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerProfileDetailsWindow.ClickOnOkBtnPopUp();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        #endregion

        #region Garnishments Tab Tests

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerTransactionHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerGarnishment);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[8].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(4);
                    WorkerGarnishmentsWindow.SelectTransactionHistoryTab();
                    Playback.Wait(2000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Garnishment Select Transaction History Tab not Displayed");
                    WorkerGarnishmentsWindow.EnterDataInTransactionHistoryTab(datarow);
                    Playback.Wait(2000);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerExistingOrdersTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerGarnishment);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[8].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(4);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Garnishment Existing Orders Tab not Displayed");
                    WorkerGarnishmentsWindow.SelectDataInComboBox(datarow);
                    Playback.Wait(2000);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        #endregion

        #region Skills Tab Tests

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerJobSkillsTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(5);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Skills Tab not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void AddSkillsForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(5);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Skills Tab not Displayed");
                    WorkerSkillsWindow.ClickOnAddorUpdateBtn();
                    Factory.AssertIsTrue(WorkerSkillsWindow.VerifyAddWorkerSkillsWindowDisplayed(),
                        "Add Worker Skills Window Not Displayed");
                    WorkerSkillsWindow.EnterDataAddWorkerSkills(datarow);
                    Playback.Wait(1000);
                    WorkerSkillsWindow.ClickOnSaveBtn();
                    WorkerSkillsWindow.SelectDutyChkBox();
                    WorkerSkillsWindow.SelectVehicleChkBox();
                    WorkerSkillsWindow.ClickonSaveBtninSkillsTab();
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerSkillsWindow.ClickonOkBtn();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void AddLicensesForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(5);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Skills Tab not Displayed");
                    WorkerSkillsWindow.EnterLicenseData(datarow);
                    WorkerSkillsWindow.SelectDutyChkBox();
                    WorkerSkillsWindow.SelectVehicleChkBox();
                    WorkerSkillsWindow.ClickonSaveBtninSkillsTab();
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerProfileDetailsWindow.VerifyWorkerPopUpDisplayed(),
                        "Confirmation Window Not Displayed");
                    WorkerSkillsWindow.ClickonOkBtn();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        #endregion

        #region Work and Payment History Tab Tests

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerWorkHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerPayment);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[9].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(6);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Work and Payment History Tab not Displayed");
                    WorkerWorkandPaymentHistoryWindow.EnterDataInWorkerHistoryTab(datarow);
                    Playback.Wait(1000);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerPaymentHistoryTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerPayment);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[9].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(6);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Work and Payment History Tab not Displayed");
                    WorkerWorkandPaymentHistoryWindow.EnterDataInPaymentHistoryTab(datarow);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        #endregion

        #region Ratings and Notes Tab Tests

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerRatingsandNotesTab()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[10].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(7);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Ratings and Notes Tab not Displayed");
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }


            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void AddRatingsForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[10].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(7);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Ratings and Notes Tab not Displayed");
                    WorkerRatingsandNotesWindow.SelectRatingsInComboBox();
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerRatingsandNotesWindow.VerifyRatingsWindowDisplayed(),
                        "Ratings Window Not Displayed");
                    WorkerRatingsandNotesWindow.EnterdataInRatingsWindow(datarow);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void AddNotesForWorker()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[10].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(7);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Workers Ratings and Notes Tab not Displayed");
                    WorkerRatingsandNotesWindow.SelectNotesInComboBox();
                    WorkerRatingsandNotesWindow.EnterdataInNotesWindow(datarow);
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void BrowseForCustomerInRatingsWindow()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[10].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(7);
                    WorkerRatingsandNotesWindow.SelectRatingsInComboBox();
                    WorkerRatingsandNotesWindow.ClickOnCustomerBrowseBtn();
                    Factory.AssertIsTrue(WorkerRatingsandNotesWindow.VerifyCustomerSearchWindowDisplayed(),
                        "Customer Search Window Not Displayed");
                    WorkerRatingsandNotesWindow.EnterDataCustomerSearchWindow(datarow);
                    WorkerRatingsandNotesWindow.ClickCancelRatings();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void BrowseForCustomerInNotesWindow()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Workers");
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerRating);
            foreach (var datarow in datarows)
            {
                var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[10].ToString());
                if (worker)
                {
                    Playback.Wait(2000);
                    WorkerSummaryWindow.SelectMainTab(7);
                    WorkerRatingsandNotesWindow.SelectNotesInComboBox();
                    WorkerRatingsandNotesWindow.ClickOnBrowseBtn();
                    Factory.AssertIsTrue(WorkerRatingsandNotesWindow.VerifyCustomerSearchWindowDisplayed(),
                       "Customer Search Window Not Displayed");
                    WorkerRatingsandNotesWindow.EnterDataCustomerSearchWindow(datarow);
                    WorkerRatingsandNotesWindow.ClickCancelnotes();
                    WorkerSummaryWindow.ClickOnCloseBtn();
                }
                Factory.AssertIsTrue(worker, "Requested Worker not found");
            }
            Cleanup();
        }

        #endregion

        #region Survey Tab Tests

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void VerifyWorkerSurveyTab()
        {
            try
            {
                Initialize();
                LandingPage.SelectFromToolbar("Workers");
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerSkills);
                foreach (var datarow in datarows.Where(
                    x => !(string.IsNullOrWhiteSpace(x.ItemArray[17].ToString()))))
                {
                    var worker = WorkerSummaryWindow.SelectWorkerFromTable(datarow.ItemArray[17].ToString());
                    if (worker)
                    {
                        Playback.Wait(2000);
                        WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Survey);
                        Playback.Wait(1000);
                        Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                            "Workers Survey Tab not Displayed");
                        WorkerSummaryWindow.ClickOnCloseBtn();
                    }
                    Factory.AssertIsTrue(worker, "Requested Worker not found");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Worker"), TestCategory("Positive")]
        public void RaChangeWorkerStatus()
        {
            try
            {
                WindowsActions.KillEllisProcesses();
                App = EllisHome.LaunchEllisAsRAUser();
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAdvancedSearch);
                foreach (var datarow in datarows)
                {
                    SearchWindow.SelectSearchElements(null, "Worker", SearchWindow.SearchTypeConstants.Advanced);
                    WorkerAdvancedSearchWindow.EnterWorkerNameAsSearchData("TEST");
                    var worker = WorkerAdvancedSearchWindow.SelectWorkerfromSearchResults(datarow.ItemArray[36].ToString());
                    Factory.AssertIsTrue(worker, "Requested Worker not found");
                    if (!worker) continue;

                    Playback.Wait(2000);
                    WorkerAdvancedSearchWindow.SelectApplicationChkBox(datarow);
                    WorkerSummaryWindow.SelectMainTab(WorkerSummaryWindow.WorkerProfileTabConstants.Survey);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(), "Workers Survey not Displayed");
                    WorkerSurveyWindow.EnterDatainSsn(datarow);
                    WorkerSurveyWindow.EnterDataInSurveyGrid();
                    WorkerSurveyWindow.EnterNotes(datarow);
                    Playback.Wait(1000);
                    Factory.AssertIsTrue(WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed(),
                        "Worker Profile Window Not Displayed");
                }
            }
            finally
            {
                Cleanup();
            }

        }

        #endregion

        #region CleanUp Test

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

        #endregion
    }
}