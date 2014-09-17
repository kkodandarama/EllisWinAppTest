using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EllisWinAppTest.Windows.EllisWindow;
using System.Diagnostics;

namespace EllisWinAppTest.SearchTests
{
    [CodedUITest]
    public class SimpleSearchTest : AppContext
    {
        public IEnumerable<DataRow> Initialize(int retries = 5)
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            if (!App.WaitForControlReady(6000) && retries >= 0)
                return Initialize(retries - 1);

            return ExcelReader.ImportSpreadsheet(ExcelFileNames.SimpleSearch);
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void CustomerSimpleSearch()
        {
            try
            {
                var datarows = Initialize();
                foreach (var dataRow in datarows.Where(dataRow => dataRow.ItemArray[4].ToString().Equals("Customer")))
                {
                    Debug.WriteLine(dataRow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(dataRow.ItemArray[5].ToString(), "Customer",
                        SearchWindow.SearchTypeConstants.Simple);
                    Playback.Wait(3000);
                    Globals.CustomerName = dataRow.ItemArray[6].ToString();

                    Factory.AssertIsTrue(SimpleSearchWindow.VerifyCustomerProfileDisplayed(dataRow.ItemArray[6].ToString()),
                        "Customer profile for " + dataRow.ItemArray[5] + " Is not displayed");

                    TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetCustomerProfileWindowProperties());
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void WorkerSimpleSearch()
        {
            var datarows = Initialize();
            foreach (var dataRow in datarows)
            {
                var TypeOne = dataRow.ItemArray[4].ToString();

                if (TypeOne.Equals("Worker"))
                {
                    Console.WriteLine(dataRow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(dataRow.ItemArray[5].ToString(), TypeOne,
                        SearchWindow.SearchTypeConstants.Simple);
                    Globals.WorkerName = dataRow.ItemArray[6].ToString();

                    SimpleSearchWindow.SelectWorkerFromResultsWindow();

                    Factory.AssertIsTrue(SimpleSearchWindow.VerifyWorkerProfileDisplayed(dataRow.ItemArray[6].ToString()), "Displayed results does not contain: " + dataRow.ItemArray[5]);
                    
                    TitlebarActions.ClickClose((WinWindow) SimpleSearchWindow.GetWorkerProfileWindowProperties());
                    TitlebarActions.ClickClose((WinWindow) SimpleSearchWindow.GetSearchResultsWindowProperties());
                }
            }
            //Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void NotificationSimpleSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("Lockout")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Lockout",
                    SearchWindow.SearchTypeConstants.Simple);
                Playback.Wait(3000);

                Factory.AssertIsTrue(SimpleSearchWindow.VerifyLockoutOverideWindowDisplayed(), "Results are displayed for the search criteria");
                TitlebarActions.ClickClose((WinWindow) SimpleSearchWindow.GetLockoutOverideWindowProperties());
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void BillingLineItemSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("BillingLineItem")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "BillingLineItem",
                    SearchWindow.SearchTypeConstants.Simple);
                Playback.Wait(3000);

                Factory.AssertIsTrue(SimpleSearchWindow.VerifySearchResultsWindowDisplayed(), "Search results window is not displayed for the search criteria");
                SimpleSearchWindow.CloseResultsWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void InvoicesSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("Invoices")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Invoices2",
                    SearchWindow.SearchTypeConstants.Simple);
                Playback.Wait(3000);

                Factory.AssertIsTrue(SimpleSearchWindow.VerifyCustomerInvoiceDisplayed(), "Customer invoice is not displayed for the search criteria");
                    
                TitlebarActions.ClickClose((WinWindow) SimpleSearchWindow.GetCustomerInvoiceWindowProperties());
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void CreditCardSearch()
        {
            try
            {
                var datarows = Initialize();
                foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("CreditCard")))
                {
                    Debug.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "CreditCard",
                        SearchWindow.SearchTypeConstants.Simple);
                    SearchWindow.SelectCardFromResults();

                    Factory.AssertIsTrue(SimpleSearchWindow.VerifyViewCreditCardPaymentWindowDisplayed(datarow.ItemArray[5].ToString()), 
                        "View Credit card payment window is not displayed for the search criteria");

                    TitlebarActions.ClickClose((WinWindow)SimpleSearchWindow.GetViewCreditCardPaymentWindowProperties());
                    SimpleSearchWindow.CloseResultsWindow();
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void InvoiceTransactionsSearch()
        {
            try
            {
                var datarows = Initialize();
                foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("ITransactions")))
                {
                    Console.WriteLine(datarow.ItemArray[3]);
                    SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "ITransactions2",
                        SearchWindow.SearchTypeConstants.Simple);
                    SearchWindow.SelectInvoiceFromResults();

                    Factory.AssertIsTrue(SimpleSearchWindow.VerifyCustomerInvoiceWindowDisplayed(datarow.ItemArray[5].ToString()), "View Credit card payment window is not displayed for the search criteria");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void InvoiceRelationShipsSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("IRelationShips")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "IRelationShips2",
                    SearchWindow.SearchTypeConstants.Simple);
                Factory.AssertIsTrue(SimpleSearchWindow.VerifySearchResultsWindowDisplayed(), "Results are displayed for the search criteria");
                SimpleSearchWindow.CloseResultsWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void JobOrderSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("JobOrder")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "JobOrder",
                    SearchWindow.SearchTypeConstants.Simple);
                Factory.AssertIsTrue(OpenJobOrder.VerifyJobOrderWindowDisplayed(datarow.ItemArray[5].ToString()), "Job order window is not displayed for the search criteria");
                OpenJobOrder.CloseJobOrderProfile();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void QuoteSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("Quote")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Quote",
                    SearchWindow.SearchTypeConstants.Simple);

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteProfileWindowDisplayed(), "Quote profile window is not displayed");
                CustomerProfileWindow.CloseQuoteProfileWindow();
                Playback.Wait(5000);
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void DispatchSearch()
        {
            var datarows = Initialize();

            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("Dispatch")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "Dispatch",
                    SearchWindow.SearchTypeConstants.Simple);
                SearchWindow.SelectJobOrderFromResults();

                Factory.AssertIsTrue(JobOrderWindow.VerifyDispatchStatusDisplayed(), "Dispatch status is not displayed for selected job order");
                JobOrderWindow.CloseJobOrderProfileWindow();
                SimpleSearchWindow.CloseResultsWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void WorkTicketSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("WorkTicket")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "WorkTicket",
                    SearchWindow.SearchTypeConstants.Simple);

                Factory.AssertIsTrue(JobOrderWindow.VerifyDispatchPayoutWindowDisplayed("REPUBLIC SERVICES / DIV 4183 - Job Order Schedule – Dispatch / Payout"), "Results are displayed for the search criteria");
                TitlebarActions.ClickClose((WinWindow)JobOrderWindow.GetJobOrderDispatchPayoutWindowProperties("REPUBLIC SERVICES / DIV 4183 - Job Order Schedule – Dispatch / Payout"));
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Simple Search"), TestCategory("Positive")]
        public void CheckRegisterSearch()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[4].ToString().Equals("CheckRegister")))
            {
                Console.WriteLine(datarow.ItemArray[3]);
                SearchWindow.SelectSearchElements(datarow.ItemArray[5].ToString(), "CheckRegister",
                    SearchWindow.SearchTypeConstants.Simple);
                Playback.Wait(3000);

                Factory.AssertIsTrue(SimpleSearchWindow.VerifySearchResultsWindowDisplayed(), "Search results window is not displayed for the search criteria");
                SimpleSearchWindow.CloseResultsWindow();
                Playback.Wait(5000);
            }
            Cleanup();
        }

        //[Ignore]
        //[TestMethod]
        //public void CompleteSimpleSearch()
        //{
        //    var datarows = Initialize();
        //    foreach (var dataRow in datarows)
        //    {
        //        var TypeOne = dataRow.ItemArray[4].ToString();
        //        Console.WriteLine(dataRow.ItemArray[3]);
        //        SearchWindow.SelectSearchElements(dataRow.ItemArray[5].ToString(), TypeOne,
        //                SearchWindow.SearchTypeConstants.Simple);

        //        if (TypeOne.Equals("Customer"))
        //        {
        //            Globals.CustomerName = dataRow.ItemArray[6].ToString();

        //            Factory.AssertIsTrue(SimpleSearchWindow.VerifyDisplayedResults(Globals.CustomerName),
        //                "Displayed results does not contain: " + dataRow.ItemArray[5]);
        //            CustomerProfile.CloseCustomerProfile();
        //        }

        //        if (TypeOne.Equals("Worker"))
        //        {
        //            Globals.WorkerName = dataRow.ItemArray[6].ToString();

        //            WorkerProfile.SelectWorkerFromResultsWindow();

        //            //Factory.AssertIsTrue(SimpleSearchWindow.VerifyWorkerDisplayedResults(Globals.WorkerName), "Displayed results does not contain: " + dataRow.ItemArray[5]);
        //            //WorkerProfile.CloseWorkerProfile();
        //            //WorkerProfile.CloseResultsWindow();
        //        }

        //        if (TypeOne.Equals("Lockout"))
        //        {
        //            //Factory.AssertIsTrue(SimpleSearchWindow.NullResultsDisplayed(), "Results are displayed for the search criteria");
        //            //SimpleSearchWindow.ClickRefineSearchClose();
        //            SimpleSearchWindow.CloseResultsWindow();
        //        }
        //    }

        //    Cleanup();
        //}

        [TestCleanup]
        public void Cleanup()
        {
            App.Close();
        }
    }
}