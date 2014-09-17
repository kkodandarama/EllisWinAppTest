using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerSummaryWindow : AppContext
    {

        #region Window Properties


        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        private static UITestControl GetAlertPopUpProperties()
        {
            var alertPopUp = App.Container.SearchFor<WinWindow>(new {Name = "Alert"});
            return alertPopUp;
        }

        #endregion

        #region Summary Tab Methods

        public static bool VerifyWorkersDisplayed()
        {
            var workSpace = EllisWindow.Container.SearchFor<WinWindow>(new { Name = "_layoutWorkspace" });
            var work = workSpace.Container.SearchFor<WinButton>(new { Name = "Workers" });
            return work.Enabled;
        }

        public static bool SelectWorkerFromTable(string workername)
        {
            var status = TableActions.OpenRecordFromTable(EllisWindow, "grdDispatchedWorkers", "Name", workername);
            return status;
        }

        public static void ClickOnCloseBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                MouseActions.ClickButton(workerProfileWindow, WorkerSummaryTabConstants.CloseBtn);
                
            }
           
        }

        public static bool ClickOnPrintReportBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                //var printReportBtn = Actions.GetWindowChild(workerProfileWindow, WorkerSummaryTabConstants.ReportBtn);
                //MouseActions.Click(printReportBtn);
                MouseActions.ClickButton(workerProfileWindow, WorkerSummaryTabConstants.ReportBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnChangeStatusBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                //var changeStatusBtn = Actions.GetWindowChild(workerProfileWindow, WorkerSummaryTabConstants.ChangeStatusBtn);
                //MouseActions.Click(changeStatusBtn);
                MouseActions.ClickButton(workerProfileWindow, WorkerSummaryTabConstants.ChangeStatusBtn);
                return true;
            }
            return false;
        }

        public static bool CloseAlertPopUp()
        {
            var alertPopUp = GetAlertPopUpProperties();
            if (alertPopUp.Exists)
            {
                MouseActions.ClickButton(alertPopUp, "_OKButton");
                return true;
            }
            return false;
        }

        public static bool VerifyAlertPopUpDisplayed()
        {
            var alertPopUp = GetAlertPopUpProperties();
            return (alertPopUp.Exists);
        }

        public static bool CancelAlertPopUp()
        {
            var alertPopUp = GetAlertPopUpProperties();
            if (alertPopUp.Exists)
            {
                MouseActions.ClickButton(alertPopUp, "_CancelButton");
                return true;
            }
            return false;
        }

        public static void ClosePrintWindow()
        {
            SendKeys.SendWait("{ESC}");
        }

        public static void SelectMainTab(string tabName)
        {
            
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var tabs = Actions.GetWindowChild(workerProfileWindow, "ultraTabControl2");
            var selectedMainTab = tabs.Container.SearchFor<WinTabPage>(new { Name = tabName });
            MouseActions.Click(selectedMainTab);
        }

        public static void SelectMainTab(int tabNo)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var tabs = Actions.GetWindowChild(workerProfileWindow, "ultraTabControl2");
            var selectedMainTab = tabs.Container.SearchFor<WinTabPage>(new { Name = "Summary" });
            selectedMainTab.SetFocus();

            for (int i = 1; i < tabNo; i++)
            {
                Playback.Wait(2000);
                SendKeys.SendWait("{RIGHT}");
            }
        }

        public static void SelectInnerTab(string mainTab, string innerTab)
        {
            var tabs = GetWorkerProfileTabs();
            var tabList = tabs.Container.SearchFor<WinWindow>(new { ControlName = "ultraTabControl2" });
            var wProfileTab = tabList.Container.SearchFor<WinTabPage>(new { Name = mainTab });
            var clicktInnerTab = wProfileTab.Container.SearchFor<WinTabPage>(new { Name = innerTab });
            MouseActions.Click(clicktInnerTab);
        }

        private static UITestControl GetWorkerProfileTabs()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            var children = workerProfileWindow.GetChildren();
            return children[3];
        }

        #endregion

        #region Controls

        public class WorkerProfileTabConstants
        {
            public const string Summary = "Summary";
            public const string ProfileDetails = "Profile Details";
            public const string Withholdings = "Withholdings";
            public const string Garnishments = "Garnishments";
            public const string Skills = "Skills";
            public const string WPHistory = "Work and Payment History";
            public const string RandN = "Ratings and Notes";
            public const string Survey = "Survey";
        }

        public class WorkerSummaryTabConstants
        {
            public const string RecentWorkHistoryGrid = "grdWork";
            public const string CloseBtn = "btnClose";
            public const string ChangeStatusBtn = "btnChangeStatus";
            public const string ReportBtn = "btnReport";
        }

        #endregion
    }
}