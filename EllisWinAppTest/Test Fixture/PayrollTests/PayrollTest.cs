using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.PayrollWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace EllisWinAppTest.PayrollTests
{
    [CodedUITest]
    public class PayrollTest : AppContext
    {
        [TestInitialize]
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void PayrollOCRTest()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnFilePayrollOCR();

            Factory.AssertIsTrue(LandingPage.VerifyFileDialogDisplayed(), "File dialog is not displayed");
            LandingPage.CloseFileDialogWindow();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void PayrollCreateOrderTest()
        {
            EllisHome.LaunchEllisAsPRMUser();


            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateGarnishmentOrder);
            foreach (var datarow in datarows)
            {
                EllisHome.ClickOnFilePayrollCreateOrder();
                CreateNewOrderWindow.EnterDataInCreateOrder(datarow);

                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyCreatNeworderWindowDisplayed(), "Create New Garnishment Order Window Not Displayed");
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void AddNotesInGarnishmentOrder()
        {

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.SearchGarnishmentOrder);
            foreach (var datarow in datarows)
            {
                EllisHome.LaunchEllisAsPRMUser();
                EllisHome.ClickOnFilePayrollCreateOrder();
                CreateNewOrderWindow.ClickSearchOrderButton();
                Playback.Wait(3000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifySearchOrderWindowDisplayed(), "Search Order Window Not Displayed");
                CreateNewOrderWindow.EnterDataInSearchOrderWindow(datarow);
                Playback.Wait(2000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyOrderDetailsWindowDisplayed(), "Order Details Window Not Displayed");
                CreateNewOrderWindow.SelectNotesTab();
                CreateNewOrderWindow.ClickOnAddNotesBtn();
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyAddNOteWindowDisplayed(), "Add Note Window Not Displayed");
                CreateNewOrderWindow.EnterDatainAddNoteWindow(datarow);
                CreateNewOrderWindow.CloseOrderDetailsWindow();
                CreateNewOrderWindow.ClickCloseBtnSearchOrderWindow();
                CreateNewOrderWindow.CloseCreateNewOrderWindow();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void SearchAgencyFromCreateOrder()
        {

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateGarnishmentOrder);
            foreach (var datarow in datarows)
            {
                EllisHome.LaunchEllisAsPRMUser();
                EllisHome.ClickOnFilePayrollCreateOrder();
                CreateNewOrderWindow.EnterAgencySearchData(datarow);
                CreateNewOrderWindow.ClickSearchAgencyButton();
                Playback.Wait(2000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifySearchAgencyWindowDisplayed(), "Search Agency Window Not Displayed");
                CreateNewOrderWindow.SelectAgencyRecord(datarow);
                Playback.Wait(2000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyUpdateAgencyWindowDisplayed(), "Update Agency Window Not Displayed");
                CreateNewOrderWindow.CloseUpdateAgencyWindow();
                CreateNewOrderWindow.CloseSearchAgencyWindow();
                CreateNewOrderWindow.CloseCreateNewOrderWindow();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyAnswerOrPrintLetterWindow()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnVieworPrintAnswerLetters();
            Playback.Wait(2000);

            Factory.AssertIsTrue(CreateNewOrderWindow.VerifyAnswerLetterWindowDisplayed(), "Answer Letter Window not displayed");
            CreateNewOrderWindow.CloseAnswerLetterWindow();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyRateSheetWindow()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnRateSheet();
            Playback.Wait(2000);

            Factory.AssertIsTrue(CreateNewOrderWindow.VerifyRateSheetWindowDisplayed(), "Rate Sheet Window not displayed");
            CreateNewOrderWindow.CloseRateSheetWindow();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void UpdateDateInRateSheetWindow()
        {
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateGarnishmentOrder);
            foreach (var datarow in datarows)
            {
                EllisHome.LaunchEllisAsPRMUser();
                EllisHome.ClickOnRateSheet();
                Playback.Wait(2000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyRateSheetWindowDisplayed(), "Rate Sheet Window not displayed");
                CreateNewOrderWindow.EnterRateSheetData(datarow);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyRateSheetDateWindowDisplayed(), "Rate Sheet Date Window Not Displayed");
                CreateNewOrderWindow.ClickOkinRateSheetDateWindow();
                CreateNewOrderWindow.CloseRateSheetWindow();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyGarnishmentExceptionBatchWindow()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnOrders();
            Playback.Wait(2000);
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPWUserUnprocessedJobs()
        {
            EllisHome.LaunchEllisAsPWUser();
            EllisHome.OpenPrevailingJobWindow("Unprocessed Jobs");
            Playback.Wait(5000);
            Factory.AssertIsTrue(CreateNewOrderWindow.VerifySummaryViewTitleValue("Un-Processed Job Order Payroll Reports"), "Un-Processed Job Order not Displayed");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPWUserProcessedJobs()
        {
            EllisHome.LaunchEllisAsPWUser();
            EllisHome.OpenPrevailingJobWindow("Processed Jobs");
            Playback.Wait(5000);
            Factory.AssertIsTrue(CreateNewOrderWindow.VerifySummaryViewTitleValue("Processed Job Order Payroll Reports"), "Processed Job Order not Displayed");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPrintInUnclaimedProperty()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnUnClaimedProperty();
            Playback.Wait(3000);
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.UnclaimedProperty);
            foreach (var datarow in datarows)
            {
                UnclaimedProperty.EnterDataInPrintLetterTab(datarow);
                UnclaimedProperty.ClickOnPrintLetterBtnPrintLetter();
                Factory.AssertIsTrue(UnclaimedProperty.VerifyPrintWindowDisplayed(), "Print Window Not Displayed");
                UnclaimedProperty.ClosePrintWindow();
                UnclaimedProperty.CloseUnclaimedPropertyWindow();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPrintBtnPrintLetterWindow()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnVieworPrintAnswerLetters();
            Playback.Wait(2000);
            CreateNewOrderWindow.ClickonSelectAllBtnAnswerLetter();
            Playback.Wait(2000);
            CreateNewOrderWindow.ClickonPrintBtnAnswerLetter();
            Playback.Wait(5000);
            Factory.AssertIsTrue(CreateNewOrderWindow.VerifyPrintWindowDisplayed(), "Print Window not displayed");
            CreateNewOrderWindow.ClosePrintWindow();
            CreateNewOrderWindow.CloseAnswerLetterWindow();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyOverTimeRules()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnOvertime();
            Playback.Wait(5000);
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.UnclaimedProperty);
            foreach (var datarow in datarows)
            {
                CreateNewOrderWindow.EnterDatainOverTimeWindow(datarow);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyOverTimeWindowDisplayed(), "OVerTime Window Not displayed");
                CreateNewOrderWindow.ClickModifyBtnOverTime();
                CreateNewOrderWindow.ClickokBtnSaveSucccess();
            }


        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyStateHoliday()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnStateHoliday();
            Playback.Wait(3000);
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.UnclaimedProperty);
            foreach (var datarow in datarows)
            {
                CreateNewOrderWindow.EnterDataInStateHoliday(datarow);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifyNoResultsFoundWindowDisplayed(),
                    "No Results Found Window Not Displayed");
                CreateNewOrderWindow.ClickOnOkBtnNoresults();
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPaymentRefund()
        {
            EllisHome.LaunchEllisAsGSUser();
            EllisHome.ClickOnPayment();
            Playback.Wait(3000);
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.UnclaimedProperty);
            foreach (var datarow in datarows)
            {
                CreateNewOrderWindow.EnterDataInPaymentRefundStatusWindow(datarow);

            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyPayrollAdjustments()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnPayrollAdjustments();
            Playback.Wait(3000);

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyResetBtninSearchPayments()
        {
            EllisHome.LaunchEllisAsPRMUser();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.PaymentSearch);
            foreach (var datarow in datarows)
            {
                EllisHome.ClickOnSearchPayments();
                Playback.Wait(3000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifySearchPaymentsWindowDisplayed(), "Search Payment Window Not Displayed");
                CreateNewOrderWindow.EnterDataInSearchPayment(datarow);
                CreateNewOrderWindow.ClickOnResetBtnPaymentSearch();
                Playback.Wait(5000);
                CreateNewOrderWindow.CloseSearchPaymentsWindow();
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifySearchPayments()
        {
            EllisHome.LaunchEllisAsPRMUser();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.PaymentSearch);
            foreach (var datarow in datarows)
            {
                EllisHome.ClickOnSearchPayments();
                Playback.Wait(3000);
                Factory.AssertIsTrue(CreateNewOrderWindow.VerifySearchPaymentsWindowDisplayed(), "Search Payment Window Not Displayed");
                CreateNewOrderWindow.EnterDataInSearchPayment(datarow);
                CreateNewOrderWindow.ClickOnSearchBtnPaymentSearch();
                Playback.Wait(5000);
                CreateNewOrderWindow.CloseSearchPaymentsWindow();
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifyClearedChecksGateway()
        {
            EllisHome.LaunchEllisAsGLUser();
            EllisHome.ClickOnClearedChecksGateway();
            Playback.Wait(3000);
            Factory.AssertIsTrue(CreateNewOrderWindow.VerifyClearedCheckGatewayWindowDisplayed(), "Cleared Checks Gateway Window Not Displayed");
            CreateNewOrderWindow.CloseClearedCheckGatewayWindow();

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Payroll"), TestCategory("Positive")]
        public void VerifySchedulePayments()
        {
            EllisHome.LaunchEllisAsPRMUser();
            EllisHome.ClickOnSchedulePayments();
            Playback.Wait(3000);
            Factory.AssertIsTrue(CreateNewOrderWindow.VerifySchedulePaymentWindowDisplayed(), "Schedule Payment Window Not Displayed");
            CreateNewOrderWindow.CloseSchedulePaymentWindow();

        }

        [TestCleanup]
        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}
