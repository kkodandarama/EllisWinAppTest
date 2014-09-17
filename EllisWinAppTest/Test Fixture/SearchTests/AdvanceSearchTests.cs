using System;
using System.Linq;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.DispatchAndPayoutWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.SearchWindow;
using EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using System.Diagnostics;

namespace EllisWinAppTest.SearchTests
{
    [CodedUITest]
    public class AdvanceSearchTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
            App = EllisHome.LaunchEllisAsCSRUser();
            Thread.Sleep(5000);
            //App.SetFocus();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void CustomerAdvanceSearch()
        {
            try
            {
                Initialize();
                int i = 1;
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CustomerAdvancedSearch);
                foreach (var dataRow in datarows.Where(v => !String.IsNullOrWhiteSpace(v.ItemArray[3].ToString())).Take(6))
                {
                    Debug.WriteLine("Test Case " + i++ + " - " + dataRow.ItemArray[24]);
                    SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
                    CustomerAdvanceSearchWindow.EnterAdvancedSearchData(dataRow);
                    CustomerAdvanceSearchWindow.ClickOnSearchButton();
                    Factory.AssertIsTrue(CustomerAdvanceSearchWindow.VerifySearchResultsWindowDisplayed(),
                        "Search validation error window was not displayed");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void VerifyCustomerSearchValidationErrorDisplayedTest()
        {
            Initialize();
            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.ClickOnSearchButton();
            Playback.Wait(2000);
            Factory.AssertIsTrue(CustomerAdvanceSearchWindow.VerifySearchValidationWindowDisplayed(),
                "Search validation error window was not displayed");
            Playback.Wait(2000);
            CustomerAdvanceSearchWindow.ClickValidationOk();
            CustomerAdvanceSearchWindow.CloseSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void CustomerCalendarFilterTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("Customers");
            LandingPage.ClickOnCalendarButton(LandingPage.LandingPageControls.Advanced);
            LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedFromDate, "01012014");
            LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedToDate, "02272014");
            //LandingPage.ClickOnCalendarClient();
            SendKeys.SendWait("{TAB}");
            CustomerProfileWindow.SelectFirstCustomerFromTable();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyProfileDefaults(), "");

            TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetCustomerProfileWindowProperties());

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void WorkerAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.WorkerAdvancedSearch);
            foreach (var datarow in datarows)
            {

                SearchWindow.SelectSearchElements(null, "Worker", SearchWindow.SearchTypeConstants.Advanced);
                WorkerAdvancedSearchWindow.EnterWorkerAdvancedSearchData(datarow);
                WorkerAdvancedSearchWindow.ClickSearchBtn();
                var worker = false;
                Playback.Wait(5000);
                var searchResults = WorkerAdvancedSearchWindow.GetWorkerSearchResultsWindowProperties();
                if (searchResults.Exists)
                {
                    WorkerAdvancedSearchWindow.SelectWorkerfromSearchResults(datarow.ItemArray[36].ToString());
                }
                WorkerAdvancedSearchWindow.SelectApplicationChkBox(datarow);
                if (WorkerSurveyWindow.VerifyWorkerProfileWindowDisplayed())
                {
                    WorkerSummaryWindow.ClickOnCloseBtn();
                    WorkerAdvancedSearchWindow.CloseSearchResultsWindow();
                    worker = true;
                }
                Factory.AssertIsTrue(worker, "Worker not found with given search criteria");

            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void LockoutAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.LockOutAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                SearchWindow.SelectSearchElements(null, "Lockout", SearchWindow.SearchTypeConstants.Advanced);
                NotificationAdvanceSearchWindow.EnterNotificationAdvancedSearchData(dataRow);
                NotificationAdvanceSearchWindow.ClickSearchBtn();
                Playback.Wait(3000);
                var lockOut = NotificationAdvanceSearchWindow.SelectLockOutNotification(dataRow.ItemArray[9].ToString());
                if (lockOut)
                {
                    Playback.Wait(5000);
                    NotificationAdvanceSearchWindow.ClickCancelcreditlimitBtn();
                    NotificationAdvanceSearchWindow.ClosedSearchResultsWindow();
                }
                if (NotificationAdvanceSearchWindow.VerifysearchResultsWindowDisplayed())
                    NotificationAdvanceSearchWindow.ClosedSearchResultsWindow();

            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void BillingAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Billing Search")
                {
                    SearchWindow.SelectSearchElements(null, "BillingLineItem", SearchWindow.SearchTypeConstants.Advanced);
                    ARAdvancedSearchWindow.EnterBillingSearchData(dataRow);
                    ARAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var billing = ARAdvancedSearchWindow.SelectBillingRecord(dataRow.ItemArray[10].ToString());
                    if (billing)
                    {
                        Playback.Wait(5000);
                        MouseActions.ClickButton(DispatchProfileWindow.DispatchProfileWindowProperties(), "btnCancel");

                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void CreditCardAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Credit Card")
                {
                    SearchWindow.SelectSearchElements(null, "CreditCard", SearchWindow.SearchTypeConstants.Advanced);
                    ARAdvancedSearchWindow.EnterCreditCardSearchData(dataRow);
                    ARAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var creditCard = ARAdvancedSearchWindow.SelectCreditCardRecord(dataRow.ItemArray[11].ToString());
                    if (creditCard)
                    {
                        Playback.Wait(5000);
                        MouseActions.ClickButton(SimpleSearchWindow.GetViewCreditCardPaymentWindowProperties(), "_btnClose");
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void InvoiceAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Invoice")
                {
                    SearchWindow.SelectSearchElements(null, "Invoices2", SearchWindow.SearchTypeConstants.Advanced);
                    ARAdvancedSearchWindow.EnterInvoiceSearchData(dataRow);
                    ARAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var invoice = ARAdvancedSearchWindow.SelectInvoiceRecord(dataRow.ItemArray[34].ToString() + "    ");
                    if (invoice)
                    {
                        Playback.Wait(7000);
                        MouseActions.ClickButton(CustomerProfileWindow.GetCustomerInvoiceWindowProperties(), "_btnCancel");
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }

            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void InvoiceTransactionAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Transactions")
                {
                    SearchWindow.SelectSearchElements(null, "ITransactions2", SearchWindow.SearchTypeConstants.Advanced);
                    ARAdvancedSearchWindow.EnterTransactionSearchData(dataRow);
                    ARAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var transaction = ARAdvancedSearchWindow.SelectTransactionRecord(dataRow.ItemArray[59].ToString());
                    if (transaction)
                    {
                        Playback.Wait(10000);
                        MouseActions.ClickButton(CustomerProfileWindow.GetCustomerInvoiceWindowProperties(), "_btnCancel");
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void InvoiceRelationShipsAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.ARAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Relationship")
                {
                    SearchWindow.SelectSearchElements(null, "IRelationShips2", SearchWindow.SearchTypeConstants.Advanced);
                    ARAdvancedSearchWindow.EnterRelationshipSearchData(dataRow);
                    ARAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var relationship = ARAdvancedSearchWindow.SelectInvoiceRecord(dataRow.ItemArray[75].ToString());
                    if (relationship)
                    {
                        Playback.Wait(7000);
                        MouseActions.ClickButton(CustomerProfileWindow.GetCustomerInvoiceWindowProperties(), "_btnCancel");
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void CheckRegisterAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Register")
                {
                    SearchWindow.SelectSearchElements(null, "CheckRegister", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterCheckRegisterSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var register = QOTAdvancedSearchWindow.SelectCheckRegisterRecord(dataRow.ItemArray[10].ToString());
                    if (register)
                    {
                        Playback.Wait(7000);
                        QOTAdvancedSearchWindow.ClickOkBtnPayment();
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();

                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void WorkTicketAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "WorkTicket")
                {
                    SearchWindow.SelectSearchElements(null, "WorkTicket", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterWorkTicketSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var workTicket = QOTAdvancedSearchWindow.SelectWorkTicketRecord(dataRow.ItemArray[21].ToString());
                    if (workTicket)
                    {
                        Playback.Wait(7000);
                        QOTAdvancedSearchWindow.CloseDispatchWindow();
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();

                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void QuoteAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Quote")
                {
                    SearchWindow.SelectSearchElements(null, "Quote", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterQuoteSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var quote = QOTAdvancedSearchWindow.SelectQuoteRecord(dataRow.ItemArray[40].ToString());
                    if (quote)
                    {
                        Playback.Wait(7000);
                        QOTAdvancedSearchWindow.CloseQuoteProfileWindow();
                        Playback.Wait(2000);
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();

                }
            }
            Cleanup();
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void DispatchAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "Dispatch")
                {
                    SearchWindow.SelectSearchElements(null, "Dispatch", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterDispatchSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var dispatch = QOTAdvancedSearchWindow.SelectDispatchRecord(dataRow.ItemArray[53].ToString());
                    if (dispatch)
                    {
                        Playback.Wait(7000);
                        MouseActions.ClickButton(JobOrderWindow.GetJobOrderProfileProperties(), "btnClose");
                        Playback.Wait(2000);
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();

                }
            }
            Cleanup();
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void JobOrderAdvanceSearch()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "JobOrder")
                {
                    SearchWindow.SelectSearchElements(null, "JobOrder", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterJobOrderSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    var jobOrder = QOTAdvancedSearchWindow.SelectJobOrderRecord(dataRow.ItemArray[76].ToString());
                    if (!jobOrder)
                    {
                        var winInst = JobOrderWindow.GetJobOrderProfileProperties();
                        if (winInst.Exists)
                            MouseActions.ClickButton(JobOrderWindow.GetJobOrderProfileProperties(), "btnClose");

                    }
                    if (ARAdvancedSearchWindow.VerifySearchResultsWindowDisplayed())
                        ARAdvancedSearchWindow.CloseSearchResultsWindow();
                }
            }
            Cleanup();
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void VerifyRefineSearchBtn()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "WorkTicket")
                {
                    SearchWindow.SelectSearchElements(null, "WorkTicket", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterWorkTicketSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    ARAdvancedSearchWindow.ClickOnRefineSearchBtn();
                    ARAdvancedSearchWindow.ClickCancelBtn();

                }
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void VerifyPrintBtn()
        {
            Initialize();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.QOTAdvancedSearch);
            foreach (var dataRow in datarows)
            {
                if (dataRow.ItemArray[1].ToString() == "WorkTicket")
                {
                    SearchWindow.SelectSearchElements(null, "WorkTicket", SearchWindow.SearchTypeConstants.Advanced);
                    QOTAdvancedSearchWindow.EnterWorkTicketSearchData(dataRow);
                    QOTAdvancedSearchWindow.ClickSearchBtn();
                    Playback.Wait(10000);
                    ARAdvancedSearchWindow.ClickOnPrintBtn();
                    ARAdvancedSearchWindow.CloseSearchResultsWindow();

                }
            }
            Cleanup();
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}