using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;
using System;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class ScheduleAndAdditionalChargesWindow : AppContext
    {
        private static UITestControl ScheduleAndAdditionalChargesWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            return joborderWindow;
        }

        private static UITestControl AddOrderNotesWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Order Notes" });
            return joborderWindow;
        }

        public static UITestControlCollection GetButtonColloction(UITestControl windowInstence, string butName)
        {
            var group = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
            var btnControl = group.Container.SearchFor<WinButton>(new { Name = butName });
            var btnControlcollection = Actions.GetControlCollection(btnControl);

            return btnControlcollection;
        }

        private static UITestControlCollection GetScheduleAndAdditionalChargesWindowEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ScheduleAndAdditionalChargesWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetScheduleAndAdditionalChargesWindowDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ScheduleAndAdditionalChargesWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinComboBox>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }
        
        public static void ClickOnContinueBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Continue >");

            foreach (var control in butColloction)
            {
                Mouse.Click(control);
            }
        }

        public static void ClickOnBackBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "< Back");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnSaveBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Save");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnCancelJobOrderBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Cancel Job Order");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnAddNotesBtn()
        {
            //Playback.Wait(3000);
            var windowInstence = ScheduleAndAdditionalChargesWindowProperties();
            var btn = Actions.GetWindowChild(windowInstence, "_btnAddNotes");
            Mouse.Click(btn);

        }

        public static void ClickOnOkBtn()
        {
            //Playback.Wait(3000);
            var windowInstence = AddOrderNotesWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "OK");
            foreach (var control in butColloction)
            {
                control.SetFocus();
                SendKeys.SendWait("{ENTER}");
            }
        }

        public static void EnterDataInScheduleAndAdditionalChargesWindow(DataRow data)
        {
            //var ScheduleAndAdditionalChargesWindowControlColloction = GetScheduleAndAdditionalChargesWindowEditControlCollection();
            //Factory.GetAllControlNames(ScheduleAndAdditionalChargesWindowProperties());
            JobOrderSchedule(data);

        }

        public static void JobOrderSchedule(DataRow data)
        {
            var winInst = ScheduleAndAdditionalChargesWindowProperties();
            var tableName = Actions.GetWindowChild(winInst, "_grdJobOrderSchedule");
            var table = (WinTable)tableName;
            //var tablecont = table.Rows.GetNamesOfControls();

            var row = table.Container.SearchFor<WinRow>(new { Name = "Template Add Row" });

            var cell = row.Container.SearchFor<WinCell>(new { Name = "Start Time" });
            Mouse.DoubleClick(cell);
            SendKeys.SendWait(data.ItemArray[58].ToString());

            
            var cell2 = row.Container.SearchFor<WinCell>(new { Name = "Assigned To Branch" });
            Mouse.Click(cell2);
            SendKeys.SendWait(data.ItemArray[59].ToString());

            //var cell3 = row.Container.SearchFor<WinCell>(new { Name = "MAP" });
            //Mouse.Click(cell3);
            

            var cell4 = row.Container.SearchFor<WinCell>(new { Name = "Repeat Dispatch Allowed?" });
            cell4.SetFocus();
            if (data.ItemArray[61].ToString()!="Yes")
            {
                Mouse.Click(cell4);
            }


            #region set work hrs
            var saturday = 0;

            switch (DateTime.Now.DayOfWeek)
            {
                case DayOfWeek.Saturday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 0);
                    break;

                case DayOfWeek.Sunday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 1);
                    break;

                case DayOfWeek.Monday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 2);
                    break;

                case DayOfWeek.Tuesday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 3);
                    break;

                case DayOfWeek.Wednesday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 4);
                    break;

                case DayOfWeek.Thursday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 5);
                    break;

                case DayOfWeek.Friday:
                    EnterHoursInJobOrderSchedule(data, row, saturday + 6);
                    break;
            }          

            #endregion                            
        }

        private static void EnterHoursInJobOrderSchedule(DataRow data, WinRow row, int daysAfterSaturday)
        {
            var now = DateTime.Now;
            var daysToAdd = 0;
            var cellName = String.Empty;
            var startColumn = 63;

            do
            {
                cellName = now.AddDays(daysToAdd).ToString("ddd") + DateTime.Now.AddDays(daysToAdd++).ToString("dMMM");
                var dayCell = row.Container.SearchFor<WinCell>(new { Name = cellName });
                dayCell.WaitForControlReady(6000);
                Mouse.Click(dayCell);
                //The data is in column 63 for Saturday, 65 Sunday, 67 Monday, so the pattern is start at 63, then add 2 * (days after Saturday) to get the current column
                SendKeys.SendWait(data.ItemArray[startColumn + (2 * daysAfterSaturday)].ToString());
                daysAfterSaturday++;
                SendKeys.SendWait("{TAB}");
            } while (!cellName.StartsWith("Fri"));
        }

        public static void EnterDataInJobOrderNotesWindow(DataRow data)
        {
            var scheduleAndAdditionalChargesWindowControlColloction = AddOrderNotesWindowProperties();
           // Factory.GetAllControlNames(ScheduleAndAdditionalChargesWindowControlColloction);

            var textOrderNotes = Actions.GetWindowChild(scheduleAndAdditionalChargesWindowControlColloction,
                "_txtOrderNotes");
            Actions.SetText(textOrderNotes, data.ItemArray[77].ToString());

            var dropDownOrderNotes = Actions.GetWindowChild(scheduleAndAdditionalChargesWindowControlColloction,
                "_cmbRequestedBy");
            Playback.Wait(1000);
            DropDownActions.SelectDropdownByText(dropDownOrderNotes, data.ItemArray[78].ToString());
            MouseActions.ClickButton(scheduleAndAdditionalChargesWindowControlColloction, "_btnOk");
            
        }
    }
}
