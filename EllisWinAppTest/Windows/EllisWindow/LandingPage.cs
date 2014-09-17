using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.EllisWindow
{
    public class LandingPage : AppContext
    {
        private static UITestControl GetOpenFileWindowProperties()
        {
            var ccustomerWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Open" });
            return ccustomerWindow;
        }

        public static bool VerifyJobOrderSchedulesDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var jo = workSpace.Container.SearchFor<WinButton>(new {Name = "Job Orders"});
            return jo.Enabled;
        }

        public static bool VerifyDispatchDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var dis = workSpace.Container.SearchFor<WinButton>(new { Name = "Print All Dispatch Sheets" });
            return dis.Enabled;
        }

        public static bool VerifyWorkersDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var work = workSpace.Container.SearchFor<WinButton>(new {Name = "Workers"});
            return work.Enabled;
        }

        public static bool VerifyCustomersDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var cust = workSpace.Container.SearchFor<WinButton>(new {Name = "Customers"});
            return cust.Enabled;
        }

        public static bool VerifyArDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new {Name = "_layoutWorkspace"});
            var ar = workSpace.Container.SearchFor<WinButton>(new {Name = "AR"});
            return ar.Enabled;
        }

        public static void SelectFromToolbar(string name)
        {
            var toolbar = EllisWindow.Container.SearchFor<WinToolBar>(new {Name = "ToolBar"});
            var joMenuItem = toolbar.Items[GetToolbarItem(name)];
            joMenuItem.WaitForControlEnabled();
            Mouse.Click(joMenuItem);
            if (name.Equals("Customer"))
                SelectDateRange();
            Playback.Wait(2000);
        }

        public static void ClickOnCalendarButton(string type)
        {
            var btnWindow = Actions.GetWindowChild(EllisWindow, LandingPageControls.CalanderButton);
            var button = btnWindow.Container.SearchFor<WinButton>(new {Name = type});
            Mouse.Click(button);
        }

        public static void VerifyTestCustomerPresent()
        {
            SearchWindow.SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchWindow.SearchTypeConstants.Simple);
            Globals.CustomerName = "Test_Customer7381";

            if (CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed() || 
                SearchWindow.SearchWindow.SelectFirstCustomerNameFromResults()) return;

            CustomerAdvanceSearchWindow.CloseSearchResultsWindow();
            CreateCustomerWindow.ClickOnCreateCustomer();
            CreateCustomerWindow.EnterCustomerData(null);
            CreateCustomerWindow.ClickSave();
            Playback.Wait(3000);
        }

        public static void EnterDate(string datetype, string date)
        {
            var fromWindow = Actions.GetWindowChild(EllisWindow, datetype);
            var comboBox = (WinComboBox) fromWindow;
            MouseActions.Click(comboBox);
            for (int i = 0; i < 10; i++)
            {
                SendKeys.SendWait("{BACKSPACE}");
                Playback.Wait(200);
            }
            //SendKeys.SendWait("{HOME}");
            Playback.Wait(1500);
            SendKeys.SendWait(date);
        }

        public static void SelectDateRange()
        {
            ClickOnCalendarButton(LandingPageControls.Advanced);
            var pastDate = Factory.GetPastDate().ToString();
            EnterDate(LandingPageControls.AdvancedFromDate, pastDate);
            ClickDateTextbox(LandingPageControls.AdvancedToDate);
            Playback.Wait(2000);
        }

        public static void ClickDateTextbox(string datetype)
        {
            var fromWindow = Actions.GetWindowChild(EllisWindow, datetype);
            var comboBox = (WinComboBox) fromWindow;
            MouseActions.Click(comboBox);
        }

        public static void ClickOnCalendarClient()
        {
            var clientWindow = Actions.GetWindowChild(EllisWindow, LandingPageControls.CalendarClient);
            var client = (WinClient) clientWindow;
            Mouse.Click(client);
        }

        private static int GetToolbarItem(string name)
        {
            int num = 0;
            if (name == "JobOrder")
                num = 0;
            if (name == "Dispatch")
                num = 1;
            if (name == "Workers")
                num = 2;
            if (name == "Customers")
                num = 3;
            if (name == "AR")
                num = 4;

            return num;
        }

        public static void SelectCustomerInvoicesFromNavigationExplorer()
        {
            var child = Actions.GetWindowChild(EllisWindow, LandingPageControls.NavigationExplorer);
            var btn = child.Container.SearchFor<WinButton>(new {Name = "Customer Invoices"});
            Mouse.Click(btn);
        }


        public static void SelectEODReviewFromNavigationExplorer()
        {
            var child = Actions.GetWindowChild(EllisWindow, LandingPageControls.NavigationExplorer);
            var btn = child.Container.SearchFor<WinButton>(new { Name = "End Of Day Review" });
            Mouse.Click(btn);
        }

        public static bool VerifyFileDialogDisplayed()
        {
            var window = GetOpenFileWindowProperties();
            return window.Exists;
        }

        public static void CloseFileDialogWindow()
        {
            //TitlebarActions.ClickClose((WinWindow) GetOpenFileWindowProperties());
            var openfile = GetOpenFileWindowProperties();
            openfile.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);

        }

        public class LandingPageControls
        {
            public const string CalanderButton = "btnToggle";
            public const string AdvancedFromDate = "advancedFromDate";
            public const string AdvancedToDate = "advancedToDate";
            public const string CalendarClient = "CalendarView";

            public const string Advanced = "Advanced...";
            public const string Simple = "Simple...";

            public const string NavigationExplorer = "_NavigationExplorerBar";
        }
    }
}