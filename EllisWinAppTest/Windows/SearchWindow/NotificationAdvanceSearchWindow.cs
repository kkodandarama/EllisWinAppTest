using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class NotificationAdvanceSearchWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetNotificationSearchWindowProperties()
        {
            var notificationSearchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            return notificationSearchWindow;
        }

        private static UITestControl GetlockoutWindowProperties()
        {
            var lockoutWindow =
                App.Container.SearchFor<WinWindow>(new { className = "WindowsForms10.Window.8.app.0.265601d" });
            return lockoutWindow;
        }

        private static UITestControl GetNotificationSearchResultsWindowProperties()
        {
            var searchresultsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return searchresultsWindow;
        }

        #endregion

        #region Search Results Window Methods

        public static bool VerifysearchResultsWindowDisplayed()
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            return searchresultsWindow.Exists;
        }

        public static bool SelectLockOutNotification(string lockoutnumber)
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            if (searchresultsWindow.Exists)
            {


                var label = Actions.GetWindowChild(searchresultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var lockout = TableActions.OpenRecordFromTable(searchresultsWindow, "_grdSearchResult",
                    "Lockout Request Number",
                    lockoutnumber);
                return lockout;
            }
            return false;
        }

        public static bool ClosedSearchResultsWindow()
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            if (searchresultsWindow.Exists)
            {
                 TitlebarActions.ClickClose((WinWindow) searchresultsWindow);
                return true;
            }
             return false;
        }

        public static bool ClickonRefineSearchBtn()
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            if (searchresultsWindow.Exists)
            {
                MouseActions.ClickButton(searchresultsWindow, "_btnRefineSearch");
                return true;
            }
            return false;
        }

        public static bool ClickonPrintBtn()
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            if (searchresultsWindow.Exists)
            {
                MouseActions.ClickButton(searchresultsWindow, "_btnPrint");
                return true;
            }
            return false;
        }

        public static bool ClickonExportBtn()
        {
            var searchresultsWindow = GetNotificationSearchResultsWindowProperties();
            if (searchresultsWindow.Exists)
            {
                MouseActions.ClickButton(searchresultsWindow, "_btnExport");
                return true;
            }
            return false;
        }

        #endregion

        #region Credit Limit Window Methods

        public static bool VerifylockoutWindowDisplayed()
        {
            var lockoutWindow = GetlockoutWindowProperties();
            return lockoutWindow.Enabled;
        }

        public static bool ClickCancelcreditlimitBtn()
        {
            var lockoutWindow = GetlockoutWindowProperties();
            if (lockoutWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(lockoutWindow, "_cancelButton");
                MouseActions.Click(cancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickSubmitcreditlimitBtn()
        {
            var lockoutWindow = GetlockoutWindowProperties();
            if (lockoutWindow.Exists)
            {
                var submitBtn = Actions.GetWindowChild(lockoutWindow, "_submitButton");
                MouseActions.Click(submitBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Notification Advanced Search Methods

        public static bool VerifyNotificationSearchWindowDisplayed()
        {
            var notificationSearchWindow = GetNotificationSearchWindowProperties();
            return notificationSearchWindow.Enabled;
        }

        public static void EnterNotificationAdvancedSearchData(DataRow data)
        {
            var notificationSearchWindow = GetNotificationSearchWindowProperties();

            var requestnumber = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.RequestNumber);
            Actions.SetText(requestnumber, data.ItemArray[3].ToString());

            var requestername = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.Requester);
            Actions.SetText(requestername, data.ItemArray[4].ToString());

            var customername = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.Customer);
            Actions.SetText(customername, data.ItemArray[5].ToString());

            var decisionmakername = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.DecisionMaker);
            Actions.SetText(decisionmakername, data.ItemArray[6].ToString());

            var branchname = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            DropDownActions.SelectDropdownByText(branchname, data.ItemArray[7].ToString());

            var requeststatus = Actions.GetWindowChild(notificationSearchWindow,
                NotificationSearchConstants.LockoutStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            DropDownActions.SelectDropdownByText(requeststatus, data.ItemArray[8].ToString());

        }

        public static bool ClickCancelBtn()
        {
            var notificationSearchWindow = GetNotificationSearchWindowProperties();
             if (notificationSearchWindow.Exists)
             {
                 var cancelBtn = Actions.GetWindowChild(notificationSearchWindow, NotificationSearchConstants.CancelBtn);
                 MouseActions.Click(cancelBtn);
                 return true;
             }
            return false;
        }

        public static bool ClickSearchBtn()
        {
            var notificationSearchWindow = GetNotificationSearchWindowProperties();
            if (notificationSearchWindow.Exists)
            {
                var searchBtn = Actions.GetWindowChild(notificationSearchWindow, NotificationSearchConstants.SearchBtn);
                MouseActions.Click(searchBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class NotificationSearchConstants
        {
            public const string RequestNumber = "txtLockoutRequestNumber";
            public const string DecisionMaker = "txtDMFirstName";
            public const string Customer = "txtCFirstName";
            public const string Requester = "txtRFirstName";
            public const string LockoutStatus = "cbLockoutType";
            public const string Section = "_cmbSection";
            public const string Branch = "cbBranch";
            public const string SubSection = "_cmbSubSection";
            public const string SearchBtn = "_buttonSearch";
            public const string CancelBtn = "_cancelButton";

        }

        #endregion
    }
}