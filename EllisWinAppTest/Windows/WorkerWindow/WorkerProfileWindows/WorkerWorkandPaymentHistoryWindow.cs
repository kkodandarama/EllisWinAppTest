using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerWorkandPaymentHistoryWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        #endregion

        #region Work History Tab Methods

        public static bool ClickOnPrintBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var printBtn = Actions.GetWindowChild(workerProfileWindow, WorkerWorkHistoryConstants.PrintBtn);
                Mouse.Click(printBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnGoBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var goBtn = Actions.GetWindowChild(workerProfileWindow, WorkerWorkHistoryConstants.GoBtn);
                Mouse.Click(goBtn);
                return true;
            }
            return false;
        }

        public static void EnterDataInWorkerHistoryTab(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var historyType = Actions.GetWindowChild(workerProfileWindow,
                    WorkerWorkHistoryConstants.WorkHistory);
            DropDownActions.SelectDropdownByText(historyType, "Work History");

            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                Factory.SetMaskedText(workerProfileWindow, WorkerWorkHistoryConstants.FromDate, data.ItemArray[4].ToString());
            
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                Factory.SetMaskedText(workerProfileWindow, WorkerWorkHistoryConstants.ToDate, data.ItemArray[5].ToString());
           
            MouseActions.ClickButton(workerProfileWindow, WorkerWorkHistoryConstants.GoBtn);
        }


        #endregion

        #region Payment History Tab Methods

        public static bool ClickOnPrintBtnPayment()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var printBtn = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentHistoryConstants.PrintBtn);
                Mouse.Click(printBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnGoBtnPayment()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var goBtn = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentHistoryConstants.GoBtn);
                Mouse.Click(goBtn);
                return true;
            }
            return false;
        }

        public static bool SelectPaymentHistoryInDropDown()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var historyType = Actions.GetWindowChild(workerProfileWindow,
                    WorkerWorkHistoryConstants.WorkHistory);
                historyType.SetFocus();
                SendKeys.SendWait("Payment History");
                return true;
            }
            return false;
            
        }

        public static void EnterDataInPaymentHistoryTab(DataRow data)
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var historyType = Actions.GetWindowChild(workerProfileWindow,
                    WorkerWorkHistoryConstants.WorkHistory);
            historyType.SetFocus();
            SendKeys.SendWait("Payment History");

            var paymentType = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentHistoryConstants.PaymentDetails);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
            DropDownActions.SelectDropdownByText(paymentType, data.ItemArray[6].ToString());
            
            var dateFrom = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentHistoryConstants.FromDate);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            dateFrom.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[7].ToString());

            var dateTo = Actions.GetWindowChild(workerProfileWindow, WorkerPaymentHistoryConstants.ToDate);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            dateTo.SetFocus();
            Actions.SendText(" ");
            Actions.SendText("{HOME}");
            SendKeys.SendWait(data.ItemArray[8].ToString());

            MouseActions.ClickButton(workerProfileWindow, WorkerPaymentHistoryConstants.GoBtn);
        }

        #endregion

        #region Controls

        private class WorkerWorkHistoryConstants
        {
            public const string WorkHistory = "cmbHistoryType";
            public const string FromDate = "DtEditorFromDate";
            public const string ToDate = "DtEditorTODate";
            public const string WorkHistoryGrid = "grdWork";
            public const string PrintBtn = "btnWorkPrintChanges";
            public const string GoBtn = "btnGo";
        }

        private class WorkerPaymentHistoryConstants
        {
            public const string PaymentDetails = "cmbPaymentDetails";
            public const string FromDate = "DtEditorFromDate";
            public const string ToDate = "DtEditorToDate";
            public const string PaymentHistoryGrid = "grdPayments";
            public const string PrintBtn = "btnPrint";
            public const string GoBtn = "btnGo";
        }
        #endregion
    }
}