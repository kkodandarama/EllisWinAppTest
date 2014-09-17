using System;
//using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using EllisWinAppTest.Helpers;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using System.Threading.Tasks;

namespace EllisWinAppTest.Windows.DispatchAndPayoutWindow
{
    internal class DispatchProfileWindow : AppContext
    {
        public static UITestControl DispatchProfileWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { ControlId = "DispatchAndPayoutView" });
            return winInst;
        }

        public static AutomationElement OvertimeApproachingWindowProperties()
        {
            var windows = AutomationElement.RootElement.FindAll(
                TreeScope.Descendants,
                new System.Windows.Automation.PropertyCondition(
                    AutomationElement.IsWindowPatternAvailableProperty, true));

            foreach (AutomationElement window in windows)
            {
                if (window == null || window.Current.Name == null)
                    continue;
                if (window.Current.Name.StartsWith("Over Time approaching"))
                    return window;
            }

            return null;
        }

        public static bool OvertimeWindowApproachingCancel(AutomationElement parent)
        {
            if (parent == null)
                return false;

            var cancel = parent.FindFirst(TreeScope.Descendants,
                new System.Windows.Automation.PropertyCondition(
                    AutomationElement.AutomationIdProperty, "_CancelButton"));

            object pattern;

            if (cancel.TryGetCurrentPattern(InvokePattern.Pattern, out pattern))
            {
                var buttonPattern = pattern as InvokePattern;
                if (buttonPattern != null)
                {
                    buttonPattern.Invoke();
                    return true;
                }
            }

            return false;
        }

        public static bool WorkerHoursWindow()
        {

            var windows = AutomationElement.RootElement.FindAll(
                TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(
                        AutomationElement.IsWindowPatternAvailableProperty, true));

            AutomationElement workerHours = AutomationElement.RootElement;

            foreach (AutomationElement window in windows)
            {
                if (window.Current.Name.StartsWith("Worker Hours"))
                    workerHours = window;
            }

            if (workerHours == AutomationElement.RootElement) return false;

            var yesButton = workerHours.FindFirst(TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, "btnYes"));

            object pattern;

            if (yesButton.TryGetCurrentPattern(InvokePattern.Pattern, out pattern))
            {
                var buttonPattern = pattern as InvokePattern;
                if (buttonPattern != null)
                {
                    buttonPattern.Invoke();
                    return true;
                }
            }

            return false;

        }

        public static AutomationElement PrintWindowProperties()
        {
            AutomationElement printWindow;
            var count = 0;

            do
            {
                printWindow = AutomationElement.RootElement.FindFirst(
                        TreeScope.Descendants,
                        new System.Windows.Automation.PropertyCondition(
                            AutomationElement.NameProperty, "Print"));

                Playback.Wait(500);
                ++count;
            } while (printWindow == null && count < 15);

            return printWindow;
        }

        public static void PrintWindowPropertiesClickCancel(AutomationElement printWindow)
        {
            AutomationElement cancelButton;
            var count = 0;

            do
            {
                cancelButton = printWindow.FindFirst(
                        TreeScope.Descendants,
                        new System.Windows.Automation.AndCondition(
                            new System.Windows.Automation.PropertyCondition(
                                AutomationElement.LocalizedControlTypeProperty, "button"),
                                new System.Windows.Automation.PropertyCondition(
                                AutomationElement.NameProperty, "Cancel")
                                ));
                Playback.Wait(500);
                ++count;
            } while (cancelButton == null && count < 15);

            if (cancelButton != null)
                Mouse.Click(cancelButton.GetClickablePoint());
        }

        public static UITestControl AssignWorkerWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Assign Worker(s)" });
            return winInst;
        }

        public static UITestControl SaveWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Save" });
            return winInst;
        }

        private static UITestControl PendingScheduleWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Pending Schedule" });
            return winInst;
        }

        internal static UITestControl WarningWindowProperties()
        {
            var COD = App.Container.SearchFor<WinWindow>(new { Name = "OrderDetailLanding" });
            if (COD.Exists)
                MouseActions.ClickButton(COD, "_OKButton");

            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Warning" });
            if (winInst.Exists)
                MouseActions.ClickButton(winInst, "btnOK");
            return winInst;
        }

        //Print/Email Wage Notice (Pay Stub)
        public static UITestControl PaystubWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Print/Email Wage Notice (Pay Stub)" });
            return winInst;
        }

        public static UITestControl AssignAckWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Dispatch and Payout" });
            return winInst;
        }

        public static UITestControl WorkerDispatchDetailsWindowProperties(string workerName)
        {
            var dispatchProfile1 = EllisWindow.Container.SearchFor<WinWindow>(
                                    new { Name = workerName + " - Dispatch Details" });

            return dispatchProfile1;
        }

        public static UITestControl ValidationMessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Message" });
            return winInst;
        }

        public static UITestControl ReviewDispatchWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Review Dispatch" });
            return winInst;
        }

        public static UITestControl ErrorMessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "ERROR" });
            return winInst;
        }

        public static UITestControl Error1MessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Error" });
            return winInst;
        }

        public static UITestControl Error2MessageWindowProperties()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Paycard Loading Failure" });
            return winInst;
        }

        public static bool HandleSelectWorkerValidationWindow()
        {
            var winInst = ValidationMessageWindowProperties();
            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_messageLabel");
                Console.WriteLine(control.GetProperty("Value"));
                control = Actions.GetWindowChild(winInst, "_OKButton");
                Mouse.Click(control);
                return true;
            }
            return false;
        }

        public static bool HandleWorkerNotFoundValidationWindow()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Quick Search" });
            if (winInst.Exists)
            {
                var tabResults = TableActions.SelectRecordFromTable(winInst, "grdDetails", "Worker Name", "Wiley, Sharon,A");
                MouseActions.ClickButton(winInst, "btnAdd");


                //var control = Actions.GetWindowChild(winInst, "_lblresult");
                //Console.WriteLine("Message for invalid worker search displayed as below...");
                //Console.WriteLine(control.GetProperty("Name"));
                //MouseActions.ClickButton(winInst, "btncancel");
                return true;
            }
            return false;
        }

        public static bool HandleNoWorkerNoWeekSelectedValidationWindow()
        {
            var winInst = ErrorMessageWindowProperties();
            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_messageLabel");
                Console.WriteLine("Message for not selecting worker (atleast one) and not selecting day displayed as below...");
                Console.WriteLine(control.GetProperty("Value"));
                control = Actions.GetWindowChild(winInst, "_OKButton");
                Mouse.Click(control);
                return true;
            }
            return false;
        }

        public static bool HandleEmailWorkTicketWindow(DataRow data)
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Print/Email Work Ticket" });
            //Actions.GetControlCollection();

            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_txtTicketPONum");
                Actions.SetText(control, data.ItemArray[40].ToString());

                control = Actions.GetWindowChild(winInst, "optEmail");
                Factory.SelectRadioButton(winInst, "optEmail", "Email");

                MouseActions.ClickButton(winInst, "btnCreateTicket");
                MouseActions.ClickButton(AssignWorkerWindowProperties(), "btnOK");
                Playback.Wait(5000);
                var printWindow = winInst.Container.SearchFor<WinWindow>(new { Name = "Print" });
                int i = 0;
                while (printWindow.Exists && i < 10)
                {
                    printWindow.SetFocus();
                    SendKeys.SendWait("{ESC}");
                    Playback.Wait(2000);
                    i++;
                    printWindow = App.Container.SearchFor<WinWindow>(new { Name = "Print" });
                }

                return true;
            }

            return false;
        }

        public static bool HandlePrintWorkTicketWindow(DataRow data)
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Print/Email Work Ticket" });
            //Actions.GetControlCollection();

            if (winInst.Exists)
            {
                var control = Actions.GetWindowChild(winInst, "_txtTicketPONum");
                Actions.SetText(control, data.ItemArray[40].ToString());

                //control = Actions.GetWindowChild(winInst, "optEmail");
                //Factory.SelectRadioButton(winInst, "optEmail", "Email");

                MouseActions.ClickButton(winInst, "btnCreateTicket");
                Playback.Wait(3000);

                PrintWindowPropertiesClickCancel(PrintWindowProperties());
                Playback.Wait(1500);
                PrintWindowPropertiesClickCancel(PrintWindowProperties());
                Playback.Wait(1500);

                var window = AssignWorkerWindowProperties();
                if (window.Exists)
                    MouseActions.ClickButton(window, "btnOK");
                Playback.Wait(3000);

                OvertimeWindowApproachingCancel(OvertimeApproachingWindowProperties());
                Playback.Wait(3000);

                PrintWindowPropertiesClickCancel(PrintWindowProperties());
                Playback.Wait(1500);

                var dispatchWindow = AutomationElement.RootElement.FindFirst(
                    TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(
                        AutomationElement.NameProperty, "Dispatch and Payout"));

                var OKButton = dispatchWindow.FindFirst(
                    TreeScope.Descendants,
                    new System.Windows.Automation.AndCondition(
                        new System.Windows.Automation.PropertyCondition(
                            AutomationElement.LocalizedControlTypeProperty, "button"),
                            new System.Windows.Automation.PropertyCondition(
                            AutomationElement.NameProperty, "OK")
                            ));

                Factory.AssertIsFalse(OKButton == null, "Couldn't find the 'OK' button.");
                Mouse.Click(OKButton.GetClickablePoint());

                return true;
            }

            return false;
        }

        public static bool HandleTicketNotPrinted()
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Work Ticket" });

            if (winInst.Exists)
            {
                return ClickButtonWrapper(MouseActions.ClickButton, winInst, "_OKButton");
            }

            return false;
        }

        public static bool ClickButtonWrapper(Action<UITestControl, string> act, UITestControl uit, string name)
        {
            try
            {
                act(uit, name);
            }
            catch (UITestControlNotFoundException ex)
            {
                Debug.WriteLine(String.Format("Could not find Work Ticket's OK button. \n{0}", ex));
                return false;
            }

            return true;
        }

        internal static void AddWorkerBySimpleSearch(DataRow dataRow)
        {
            var dispatchProfile = DispatchProfileWindowProperties();
            var controlInst = Actions.GetWindowChild(dispatchProfile, "txtQuickAddWorker");
            if (dataRow.ItemArray[7].ToString() != String.Empty)
            {
                Actions.SetText(controlInst, dataRow.ItemArray[7].ToString());
                MouseActions.ClickButton(dispatchProfile, "btnAdd");
            }
            else if (dataRow.ItemArray[6].ToString() != String.Empty)
            {
                Actions.SetText(controlInst, dataRow.ItemArray[6].ToString());
                MouseActions.ClickButton(dispatchProfile, "btnAdd");
            }
            else if (dataRow.ItemArray[8].ToString() != String.Empty)
            {
                Actions.SetText(controlInst, dataRow.ItemArray[8].ToString());
                MouseActions.ClickButton(dispatchProfile, "btnAdd");
            }

            //var control = Actions.GetWindowChild(winInst, "_lblresult");
            //Console.WriteLine("Message for invalid worker search displayed as below...");
            //Console.WriteLine(control.GetProperty("Name"));
            //MouseActions.ClickButton(winInst, "btncancel");


        }

        internal static void HandleWorkerSearchResultsWindow(DataRow dataRow)
        {
            var winInst = App.Container.SearchFor<WinWindow>(new { Name = "Quick Search" });
            if (winInst.Exists)
            {
                //winInst.SetFocus();
                try
                {
                    var workerName = dataRow.ItemArray[7] + ", " + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != "")
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var tabResults = Factory.SelectRecordFromTable(winInst, "grdDetails", "Worker Name", workerName);
                    if (tabResults)
                        MouseActions.ClickButton(winInst, "btnAdd");

                }
                catch
                {
                    var control = Actions.GetWindowChild(winInst, "_lblresult");
                    Console.WriteLine("Message for invalid worker search displayed as below...");
                    Console.WriteLine(control.GetProperty("Name"));
                    MouseActions.ClickButton(winInst, "btncancel");
                }
            }

            Playback.Wait(2000);
            var winInst1 = App.Container.SearchFor<WinWindow>(new { Name = "Message" });
            if (winInst1.IsControlUsable())
            {
                MouseActions.ClickButton(winInst1, "_OKButton");
            }
        }


        public static bool OpenJobOrderForDispatch(DataRow dataRow)
        {
            var calRange = Actions.GetWindowChild(EllisWindow, "btnToggle");
            if (calRange.GetProperty("Name").Equals("Advanced..."))
                Mouse.Click(calRange);
            var date1from = DateTime.Now.AddDays(-45).ToString("MM/dd/yyyy");
            var dateTo = DateTime.Now.ToString("MM/dd/yyyy");

            Factory.SetMaskedText(EllisWindow, "advancedFromDate", date1from.Replace("/", String.Empty));
            Playback.Wait(500);
            Factory.SetMaskedText(EllisWindow, "advancedToDate", dateTo.Replace("/", String.Empty));
            SendKeys.SendWait("{TAB}");
            //Job lookup takes a few moments
            Playback.Wait(6000);

            var status = LookForNonCODWarningRecord();
            //Handle Warning popup
            //var windowInst = WarningWindowProperties();

            //if (windowInst.Exists)
            //    MouseActions.ClickButton(windowInst, "btnOK");

            ////Pending Schedule
            //var windowInst1 = PendingScheduleWindowProperties();
            //if (windowInst1.Exists)
            //    MouseActions.ClickButton(windowInst1, "btnOK");

            return status;
        }

        /// <summary>
        /// If the record we select is not previously authorized, we can't use it for this test's purposes.  This function iterates through records until we find one that has previous
        /// COD authorization.
        /// </summary>
        private static bool LookForNonCODWarningRecord()
        {
            var tableName = Actions.GetWindowChild(EllisWindow, "grdDispatchJobOrder");
            var table = (WinTable)tableName;
            var rows = table.Rows;
            //We have to iterate through each Table Row until we find one that isn't Aging or COD
            for (int i = 0; i < rows.Count; ++i)
            {
                //Each row has several elements, we need to find the clickable one
                foreach (var child in rows[i].GetChildren())
                {
                    var cell = child as WinCell;
                    if (cell == null) continue;
                    if (cell.BoundingRectangle.Left == cell.BoundingRectangle.Right) continue;
                    var clickPoint = new System.Drawing.Point(cell.BoundingRectangle.Left + 3, cell.BoundingRectangle.Top + 3);
                    Mouse.Click(clickPoint);
                    Playback.Wait(200);
                    Mouse.DoubleClick(clickPoint);
                    break;
                }

                Playback.Wait(3000);
                //If the row we clicked on creates a COD window or a Aging window, we dismiss that window, then cancel the Dispatch / Payout window so we can select the next row.
                if (App.Container.SearchFor<WinWindow>(new { Name = "OrderDetailLanding" }).Exists
                    || App.Container.SearchFor<WinWindow>(new { Name = "Warning" }).Exists)
                {
                    DispatchProfileWindow.WarningWindowProperties();
                    Playback.Wait(5000);
                    AutomationElement window = AutomationElement.RootElement.FindFirst(TreeScope.Children, new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "Ellis"));

                    var windows = window.FindAll(
                        TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(
                        AutomationElement.IsWindowPatternAvailableProperty, true));

                    foreach (AutomationElement w in windows)
                    {
                        if (w.Current.Name.EndsWith("Dispatch / Payout"))
                        {
                            //cancel the dispatch / payout window so we can get back to the Ellis window and select the next row
                            Mouse.Click(w.FindFirst(TreeScope.Descendants,
                                new System.Windows.Automation.PropertyCondition(
                                    AutomationElement.AutomationIdProperty, "btnCancel"))
                                    .GetClickablePoint());

                            break;
                        }
                    }
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        //private static UITestControl GetTabs()
        //{
        //    var joProfileWindow = DispatchProfileWindowProperties();
        //    var children = joProfileWindow.GetChildren();
        //    return children[3];
        //}

        public static void SelectTab(string tabName)
        {
            var joProfileWindow = DispatchProfileWindowProperties();
            var tabList = joProfileWindow.Container.SearchFor<WinWindow>(new { ControlName = "tbcOrderSchedule" }).GetChildren();

            switch (tabName)
            {
                case "Dispatch":
                    var selectedTab = tabList[3].Container.SearchFor<WinTabPage>(new { Name = "Dispatch" });
                    MouseActions.Click(selectedTab);
                    break;

                case "Payout":
                    selectedTab = tabList[3].Container.SearchFor<WinTabPage>(new { Name = " Payout" });
                    MouseActions.Click(selectedTab);
                    break;

                case "Review Dispatch":
                    selectedTab = tabList[3].Container.SearchFor<WinTabPage>(new { Name = "Review Dispatch" });
                    MouseActions.Click(selectedTab);
                    break;
            }


        }

        public static bool SetDataInPayoutTableCell(DataRow dataRow)
        {
            for (int i = 10; i <= 22; i += 2)
            {
                if (!SetDataInPayoutTableCellHelper(dataRow, i))
                    return false;
            }

            //add customer signature
            var ddControl = Actions.GetWindowChild(DispatchProfileWindowProperties(), "cmbSignature");
            DropDownActions.SelectDropdownByText(ddControl, dataRow.ItemArray[38].ToString());

            //add customer signature notes
            var notesControl = Actions.GetWindowChild(DispatchProfileWindowProperties(), "txtNotes");
            Actions.SetText(notesControl, dataRow.ItemArray[39].ToString());

            return true;
        }

        public static bool SetDataInPayoutTableCellHelper(DataRow dataRow, int position)
        {
            var winInst = DispatchProfileWindowProperties();
            var tableName = Actions.GetWindowChild(winInst, "grdPayout");
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = "Worker" });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                if (dataRow.ItemArray[10].ToString() != "")
                    workerName = workerName + "," + dataRow.ItemArray[8];

                if (callValue == workerName)
                {
                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    rowC1.SetFocus();
                    var count = -1;
                    foreach (UITestControl control in rowC1.GetChildren())
                    {
                        if (control.Width < 1) continue;
                        count++;
                        if (count <= 5) continue;
                        if (control.Name.Contains(dataRow.ItemArray[position].ToString()))
                            ClearTableCellAndInjectText(dataRow, position, control);                            
                    }

                    return true;
                }
            }

            return false;
        }

        private static void ClearTableCellAndInjectText(DataRow dataRow, int position, UITestControl control)
        {
            if (control.Width < 1) return;
            if (control.Name.Contains(dataRow.ItemArray[position].ToString()))
            {
                ClearTableCell(control);
                try
                {
                    Keyboard.SendKeys(dataRow.ItemArray[position + 1].ToString());
                    //This TAB is to overcome a VERY STRANGE issue with the table where editing the cell causes the row to be invisible to CUIT the next time.  Tabbing out somehow changes focus? or something.
                    Keyboard.SendKeys("{TAB}");
                }
                catch 
                {
                    Debug.WriteLine("ClearTableCellAndInjectText threw an exception.");
                }
            }
        }

        private static void ClearTableCell(UITestControl control)
        {
            try
            {
                Mouse.Click(control);
                Playback.Wait(300);
                Keyboard.SendKeys("{END}");
                Playback.Wait(300);
                Keyboard.SendKeys("+{HOME}");
                Playback.Wait(300);
                Keyboard.SendKeys("{DEL}");
                Playback.Wait(300);
            }
            catch 
            {
                Debug.WriteLine("ClearTableCell threw an exception.");
            }
        }

        public static bool SelectPayoutTableColumns(DataRow dataRow)
        {
            var tableName = Actions.GetWindowChild(DispatchProfileWindowProperties(), "grdPayout");
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = "Worker" });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                if (dataRow.ItemArray[8].ToString() != "")
                    workerName = workerName + "," + dataRow.ItemArray[8];

                if (callValue == workerName)
                {

                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    rowC1.SetFocus();
                    var columnList = rowC1.GetChildren();
                    foreach (UITestControl control in columnList)
                    {
                        if (control.Width < 1) continue;
                        if (control.Name.Contains(dataRow.ItemArray[10].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[12].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[14].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[16].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[18].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[20].ToString()))
                            Mouse.Click(control);
                        if (control.Name.Contains(dataRow.ItemArray[22].ToString()))
                            Mouse.Click(control);

                    }
                    return true;
                }
            }
            return false;
        }

        public static bool SelectDispatchTableColumns(DataRow dataRow)
        {
            var tableName = Actions.GetWindowChild(DispatchProfileWindowProperties(), "grdOrderDetails");
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = "Worker" });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                if (dataRow.ItemArray[8].ToString() != "")
                    workerName = workerName + "," + dataRow.ItemArray[8];

                if (callValue == workerName)
                {

                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    rowC1.SetFocus();
                    var select = rowC1.Container.SearchFor<WinCell>(new { Name = "Select" });
                    select.SetFocus();
                    if (select.GetProperty("Value").ToString() != "True")
                    {
                        Mouse.Click(select);
                    }

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[10].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[12].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[14].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[16].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[18].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[20].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);

                    rowC1.SetFocus();
                    select = rowC1.Container.SearchFor<WinCell>(new { Name = dataRow.ItemArray[22].ToString() });
                    select.SetFocus();
                    Mouse.Click(select);
                    rowC1.SetFocus();


                    return true;
                }
            }
            return false;
        }

        public static string GetCellValueFromTable(UITestControl windowInstence, string tableControlName, string columnHeader, string columnValue, string cellHeader)
        {
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnHeader });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue == columnValue)
                {
                    rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnHeader });
                    rowHeader.SetFocus();

                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    //rowHeader.SetFocus();
                    //rowC1.SetFocus();

                    var select = rowC1.Container.SearchFor<WinCell>(new { Name = cellHeader });
                    //rowHeader.SetFocus();
                    //select.SetFocus();
                    //Mouse.Click(select);
                    callValue = select.GetProperty("Value").ToString();

                    return callValue;
                }
            }
            return "NULL";
        }

        //public static void ComapreBillRate(string dispatchDate, string billRate)
        //{
        //    var dispatchProfile1 = DispatchProfileWindowProperties();

        //    if (dispatchProfile1.Exists)
        //    {
        //        dispatchProfile1.SetFocus();
        //        var cellVal = Factory.GetCellValueFromTable(dispatchProfile1, "grdPayDetail", "Paid to Worker", "Bill Rate", dispatchDate);

        //        Console.WriteLine("Bill Rate Expected: " + billRate);
        //        Console.WriteLine("Actual Bill Rate: " + cellVal);
        //        Factory.AssertIsTrue(cellVal.ToString().Contains(billRate), "Bill Rate not matched with expected data");
        //    }
        //}

        public static WinRow GetRowFromTable(string rowName)
        {
            var dispatchProfile1 = DispatchProfileWindowProperties();
            dispatchProfile1.SetFocus();
            var row = Factory.GetRowFromTable(dispatchProfile1, "grdPayDetail", "Paid to Worker", rowName);
            return row;
        }

        public static void ComapreBillRate(WinRow row, string dispatchDate, string billRate)
        {
            var cellVal = Factory.GetCellValueFromRow(row, dispatchDate);

            Console.WriteLine("Expected Bill Rate on " + dispatchDate + " : " + billRate);
            Console.WriteLine("Actual Bill Rate on " + dispatchDate + " : " + cellVal);
            //if (!string.IsNullOrEmpty(billRate))// && cellVal != null)
            Factory.AssertIsTrue(cellVal.Equals(billRate), "Bill Rate not matched with expected data");

        }

        public static void ComaprePayRate(WinRow row, string dispatchDate, string payRate)
        {
            var cellVal = Factory.GetCellValueFromRow(row, dispatchDate);
            Console.WriteLine("Expected Pay Rate on " + dispatchDate + " : " + payRate);
            Console.WriteLine("Actual Pay Rate on " + dispatchDate + " : " + cellVal);

            Factory.AssertIsTrue(cellVal.Equals(payRate), "Pay Rate not matched with expected data");

        }

        public static void Recovery()
        {
            Assert.Fail();
            EllisHome.Initialize();
        }

        public static void Playback_PlaybackError(object sender, PlaybackErrorEventArgs e)
        {

            if (e.Error.Message.Contains("Assert failed"))
            {
                e.Result = PlaybackErrorOptions.Skip;
                Console.WriteLine("Entered If");

            }
            else
            {
                e.Result = PlaybackErrorOptions.Default;
                Console.WriteLine("Entered else");
                //EllisHome.Initialize();
            }
        }

        internal static void UnAssignUnDispatchWorker(DataRow dataRow)
        {
            var dispatchProfile = DispatchProfileWindowProperties();
            var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
            if (dataRow.ItemArray[8].ToString() != "")
                workerName = workerName + "," + dataRow.ItemArray[8];

            for (int i = 0; i <= 2; )
            {
                var workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails",
                    "Worker", workerName);
                if (workerFound)
                {
                    var setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                    Actions.SetCheckBox((WinCheckBox)setWeek, "True");
                    Playback.Wait(1000);
                    MouseActions.ClickButton(dispatchProfile, "btnUnAssign");
                    //var control = Actions.GetWindowChild(dispatchProfile, "btnUnAssign");
                    //control.SetFocus();
                    //Mouse.Click(control);
                }
                var winInst = ErrorMessageWindowProperties();
                if (winInst.Exists)
                {

                    var msgText = Actions.GetWindowChild(winInst, "_messageLabel");
                    Factory.AssertIsTrue(winInst.Exists, msgText.GetProperty("Name").ToString());
                    MouseActions.ClickButton(ErrorMessageWindowProperties(), "_OKButton");

                    break;
                }
                var dispatchPayoutWindow = AssignAckWindowProperties();
                if (dispatchPayoutWindow.Exists)
                    MouseActions.ClickButton(dispatchPayoutWindow, "_OKButton");
                i++;

            }
        }

        internal static bool RemoveDispatchWorker(DataRow dataRow)
        {
            var dispatchProfile = DispatchProfileWindowProperties();
            var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
            if (dataRow.ItemArray[8].ToString() != "")
                workerName = workerName + "," + dataRow.ItemArray[8];

            var workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);
            if (workerFound)
            {
                MouseActions.ClickButton(dispatchProfile, "btnRemoveWorker");

                var window1 = ValidationMessageWindowProperties();
                if (window1.Exists)
                    try
                    {
                        MouseActions.ClickButton(window1, "btnYes");
                    }
                    catch (Exception)
                    {

                        MouseActions.ClickButton(window1, "btnOK");
                    }


                var window2 = AssignAckWindowProperties();
                if (window2.Exists)
                    MouseActions.ClickButton(window2, "_OKButton");

                return true;
            }
            return false;
        }

        public static void HandlePrintWorkerPayoutWindow(DataRow dataRow)
        {
            var wininst = Error2MessageWindowProperties();
            if (wininst.Exists)
            {
                MouseActions.ClickButton(wininst, "_OKButton");
            }
            else
            {
                wininst = PaystubWindowProperties();
                if (wininst.Exists)
                {
                    Actions.SetCheckBox(wininst, "chkEMail0", "True");
                    //Factory.SelectRadioButton(wininst, "optEmail0", "Send to other email address:");
                    //Actions.SelectRadioButton(wininst, "Send to:");
                    var control = Actions.GetWindowChild(wininst, "otherEmail0");
                    Actions.SetText(control, dataRow.ItemArray[41].ToString());
                    MouseActions.ClickButton(wininst, "btnPrintEmail");
                    wininst = Error1MessageWindowProperties();

                    if (wininst.Exists)
                    {
                        MouseActions.ClickButton(wininst, "_OKButton");
                    }

                    wininst = SaveWindowProperties();
                    if (wininst.Exists)
                    {
                        MouseActions.ClickButton(wininst, "_OKButton");
                    }
                }
            }
        }

        internal static void HandleErrorMessageWindow()
        {
            var winInst = ErrorMessageWindowProperties();
            if (winInst.Exists)
            {
                var msgText = Actions.GetWindowChild(winInst, "_messageLabel");
                Console.WriteLine(msgText.GetProperty("Name").ToString());
                MouseActions.ClickButton(winInst, "_OKButton");
            }

        }
    }
}
