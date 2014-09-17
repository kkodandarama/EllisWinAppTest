using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using Microsoft.SqlServer.Server;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class SimpleSearchWindow : AppContext
    {

        #region windows

        public static UITestControl GetWorkerProfileWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-"+Globals.WorkerName });
            return window;
        }

        public static UITestControl GetSearchResultsWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return window;
        }

        public static UITestControl GetLockoutOverideWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "CreditLimit-Lockout Override Request" });
            return window;
        }

        public static UITestControl GetCustomerInvoiceWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Customer Invoice" });
            return window;
        }

        public static UITestControl GetViewCreditCardPaymentWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "View Credit Card Payment" });
            return window;
        }

        #endregion

        public static bool VerifyCustomerProfileDisplayed(string data)
        {
            var window = CustomerProfileWindow.GetCustomerProfileWindowProperties();
            return window.Name.Contains(data);
        }

        public static bool VerifyWorkerProfileDisplayed(string data)
        {
            var window = GetWorkerProfileWindowProperties();
            return window.Name.Contains(data);
        }

        public static bool VerifySearchResultsWindowDisplayed()
        {
            var window = GetSearchResultsWindowProperties();
            return window.Exists;
        }

        public static bool VerifyCustomerInvoiceDisplayed()
        {
            var window = GetCustomerInvoiceWindowProperties();
            return window.Exists;
        }

        public static bool VerifyViewCreditCardPaymentWindowDisplayed(string data)
        {
            var window = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, 
                new PropertyCondition(AutomationElement.NameProperty, "View Credit Card Payment"));

            var cardNum = window.FindFirst(TreeScope.Descendants, 
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "_txtCardNumber"));                 
            
            return cardNum.Current.Name.EndsWith(data);
        }
               

        public static bool VerifyCustomerInvoiceWindowDisplayed(string data)
        {
            var window = GetCustomerInvoiceWindowProperties();
            Factory.AssertIsFalse(window == null, "Couldn't find the Customer Invoice window.");
            var sw = new Stopwatch();
            sw.Start();
            var label = new WinEdit();
            do
            {
                label = window.Container.SearchFor<WinEdit>(new { ControlName = SearchControls.InvoiceNumber });
                if (label != default(WinEdit)) return true;
            } while (sw.Elapsed.Seconds < 30);

            return false;
        }

        public static void CloseResultsWindow()
        {
            try
            {
                TitlebarActions.ClickClose((WinWindow)GetSearchResultsWindowProperties());
            }
            catch (Exception)
            {
                //Suppress all exceptions
            }
        }

        public static void SelectWorkerFromResultsWindow()
        {
            var window = GetSearchResultsWindowProperties();
            var row = TableActions.SelectRowFromTable(window, SearchControls.SearchResultGrid,
                "WorkerSearchResultDomain row 1");

            var cell = row.Container.SearchFor<WinCell>(new {Value = Globals.WorkerName});
            MouseActions.DoubleClick(cell);
        }

        public static bool VerifyLockoutOverideWindowDisplayed()
        {
            var window = GetLockoutOverideWindowProperties();
            return window.Exists;
        }

        private class SearchControls
        {
            public const string SearchResultGrid = "_grdSearchResult";
            public const string InvoiceNumber ="_lblInvoiceNumber";
            public const string CardNumber = "_txtCardNumber";
        }

    }


}