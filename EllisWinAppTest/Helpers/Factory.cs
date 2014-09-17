using System.Diagnostics;
using System.IO;
using System.Windows.Automation;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Point = System.Drawing.Point;

namespace EllisWinAppTest.Helpers
{
    public class Factory : AppContext
    {
        public static void GetWinProperties()
        {
            EllisWindow = App.SearchFor<WinWindow>(new { Name = "Ellis" });
            
        }

        public static string GetPastDate()
        {
            var date = DateTime.Today.AddMonths(-5).ToString("d");
            return date;
        }

        public static void AssertIsTrue(bool condition, string isFalseMessage)
        {
            try
            {
                Assert.IsTrue(condition, isFalseMessage);
            }
            catch (Exception ex)
            {

                Debug.WriteLine("Assert Failed: " + isFalseMessage);
                Debug.WriteLine("Assert Message: " + ex.Message);
            }
        }

        public static void AssertIsFalse(bool condition, string isTrueMessage)
        {
            try
            {
                Assert.IsFalse(condition, isTrueMessage);
            }
            catch (Exception ex)
            {

                Debug.WriteLine("Assert Failed: " + isTrueMessage);
                Debug.WriteLine("Assert Message: " + ex.Message);
            }
        }
        //public static ApplicationUnderTest LaunchAppAsDifferentUser(string appPath, string username, string password,
        //    string domain)
        //{
        //    var s = new NetworkCredential("", password).SecurePassword;
        //    var path = @appPath;
        //    var myProc = new ProcessStartInfo(path);

        //    try
        //    {
        //        myProc.Domain = domain;
        //        myProc.UserName = username;
        //        myProc.Password = s;
        //        myProc.UseShellExecute = false;
        //        myProc.LoadUserProfile = true;
        //        myProc.WorkingDirectory = Path.GetDirectoryName(path);
        //    }
        //    catch (Exception)
        //    {
        //        // error handling
        //    }

        //    App = ApplicationUnderTest.Launch(myProc);
        //    return App;
        //}

        //method to click on button in any window
        // Provide window instance as 1st parameter and button name as 2nd parameter
        //public static void ClickOnButton(UITestControl windowInstence, string butName)
        //{
        //    var group = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
        //    var btnControl = group.Container.SearchFor<WinButton>(new { Name = butName });
        //    var btnControlcollection = Actions.GetControlCollection(btnControl);

        //    foreach (var control in btnControlcollection)
        //    {
        //        //MouseActions.Click(control);
        //        control.SetFocus();
        //        SendKeys.SendWait("{ENTER}");
        //    }
        //}

        public static void ClickOnButton(UITestControl windowInstence, string butName)
        {
            var winGroup = windowInstence.Container.SearchFor<WinGroup>(new { Name = "" });
            var btnControl = winGroup.Container.SearchFor<WinButton>(new { Name = butName });
            Mouse.Click(btnControl);

        }

        public static ApplicationUnderTest GetAUT()
        {
            var processes = Process.GetProcessesByName("Ellis");
            App = processes.Length > 0
                ? ApplicationUnderTest.FromProcess(processes[0])
                : ApplicationUnderTest.Launch(TestData.Path, TestData.AltPath);

            return App;
        }

        public static string GenerateNumber()
        {
            Random random = new Random();
            string r = "";
            int i;
            for (i = 1; i < 11; i++)
            {
                r += random.Next(0, 9).ToString();
            }
            return r;
        }

        public static Point GetMouseCoOrdinates(UITestControl control)
        {
            var clickPoints = new Point(control.BoundingRectangle.Width / 2 + control.BoundingRectangle.X,
                control.BoundingRectangle.Height / 2 + control.BoundingRectangle.Y);
            return clickPoints;
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

        public static void SelectDropdownByText(UITestControl dropdownControl, string text)
        {
            var control = (WinComboBox)dropdownControl;
            control.SetFocus();
            string item;

            item = control.SelectedItem;

            if (string.IsNullOrEmpty(item))
            {
                control.SetFocus();
                SendKeys.SendWait(text);
                Playback.Wait(500);
            }
        }

        public static void GetAllControlNames(UITestControl control)
        {
            var group = control.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);

            var dDownControl = group.Container.SearchFor<WinComboBox>(new { Name = "" });
            var ddownControlCollection = Actions.GetControlCollection(dDownControl);

            var btnControl = group.Container.SearchFor<WinButton>(new { Name = "" });
            var btnControlCollection = Actions.GetControlCollection(btnControl);

            //var chkboxCotrol = group.Container.SearchFor<WinCheckBox>(new { Name = "Same as above" });
            //var cheboxControlCollection = Actions.GetControlCollection(chkboxCotrol);

            var tabCotrol = group.Container.SearchFor<WinTabPage>(new { Name = "Skills" });
            var tabControlCollection = tabCotrol.FindMatchingControls();

            var chkboxCotrol = group.Container.SearchFor<WinTable>(new { Name = "" });
            var cheboxControlCollection = chkboxCotrol.FindMatchingControls();

            TextWriter tsw = new StreamWriter(@"C:\ControlCollection.txt");

            //Writing text to the file.
            tsw.WriteLine("-----------------------Text Box Controls\n\n\n");
            foreach (var edit in editControlcollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Drop Down Controls\n\n\n");
            foreach (var edit in ddownControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }


            tsw.WriteLine("Button Controls\n\n\n");
            foreach (var edit in btnControlCollection)
            {
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Table Controls\n\n\n");
            foreach (var edit in cheboxControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }

            tsw.WriteLine("-----------------------Tab Controls\n\n\n");
            foreach (var edit in tabControlCollection)
            {
                //if (edit.Enabled)
                tsw.WriteLine(edit.GetProperty("ControlName"));
            }
            //Close the file.
            tsw.Close();
        }


        //public static string GenerateNewName(string name)
        //{
        //    var n = new Random().Next(99999);
        //    var newName = name + n;
        //    return newName;
        //}

        //public static string GenerateSSNNumber()
        //{
        //    var rnd = new Random();
        //    var num = rnd.Next(100000000, 999999999).ToString();
        //    return num;
        //}

        //public static string GenerateRandomNumber(int length)
        //{
        //    var rnd = new Random();
        //    var num = rnd.Next(0, length).ToString();
        //    return num;
        //}

        //public static String GetTextBetween(String source, String leftWord, String rightWord)
        //{
        //    return
        //        Regex.Match(source, String.Format(@"{0}\s(?<words>[\w\s\/1234567890]+)\s{1}", leftWord, rightWord),
        //                    RegexOptions.IgnoreCase).Groups["words"].Value;
        //}

        //public static void SelectRadioButton(string radioButtonControl, string radioButtonGroupName)
        //{
        //    var btnCont = Actions.GetControlCollection(radioButtonControl);

        //    foreach (var control in btnCont)
        //    {
        //        if (control.GetProperty("ControlName").ToString() == radioButtonGroupName)
        //            Mouse.Click(control);
        //    }
        //}

        //public static void SelectRadioButton(WinRadioButton radioButtonControl)
        //{
        //    var btnCont = Actions.GetControlCollection(radioButtonControl);

        //    foreach (var control in btnCont)
        //    {
        //        Mouse.Click(control);
        //    }
        //}

        public static bool OpenRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        {
            int i = 0;
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            do
            {
                var rows = table.Rows;
                rows[i].SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();
                i++;

                if (callValue == columnValue)
                {
                    Mouse.Click(rowHeader);
                    Mouse.DoubleClick(rowHeader);
                    return true;
                }

            } while (i < table.Rows.Count);


            return false;
        }

        public static bool OpenFirstRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName)//, string columnValue)
        {

            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            var rows = table.Rows;
            if (rows.Count > 0)
            {
                rows[0].SetFocus();
                var rowHeaders = rows[0].GetChildren();
                rowHeaders[2].SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                //Mouse.Click(rows[0]);
                //Mouse.DoubleClick(rows[0]);
                rowHeader.SetFocus();
                Mouse.Click(rowHeader);
                Mouse.DoubleClick(rowHeader);
                return true;
            }

            return false;
        }


        public static string GetCellValueFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue, string cellHeader)
        {
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue == columnValue)
                {
                    Mouse.Click(rowHeader);
                    table.SetFocus();
                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    //Mouse.Click(rowC1);
                    //rowC1.SetFocus();
                    //Playback.Wait(1000);
                    var select = rowC1.Container.SearchFor<WinCell>(new { Name = cellHeader });
                    select.SetFocus();
                    var cellValue = select.GetProperty("Value").ToString();
                    return cellValue;
                }
            }
            return "NULL";
        }

        /// <summary>
        /// Returns the center point of a control unless that control is null or has no width, in which case it returns 0,0
        /// </summary>
        public static Point GetUITestControlMiddlePoint(UITestControl control)
        {
            if (control == null || control.BoundingRectangle.Width < 1)
                return new Point(0,0);

            return new Point(control.BoundingRectangle.Left + (int)(.5 * control.BoundingRectangle.Width),
                    control.BoundingRectangle.Top + (int)(.5 * control.BoundingRectangle.Height));
        }

        /// <summary>
        /// Returns the center point of an AutomationElement unless that control is null or has no width, in which case it returns 0,0
        /// </summary>
        public static Point GetAutomationElementMiddlePoint(AutomationElement control)
        {
            if (control == null || control.Current.BoundingRectangle.Width < 1)
                return new Point(0, 0);

            return new Point(control.Current.BoundingRectangle.Left + (int)(.5 * control.Current.BoundingRectangle.Width),
                    control.Current.BoundingRectangle.Top + (int)(.5 * control.Current.BoundingRectangle.Height));
        }

        public static WinRow GetRowFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        {
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue == columnValue)
                {
                    Mouse.Click(rowHeader);
                    table.SetFocus();
                    var rowC1 = table.Container.SearchFor<WinRow>(new { Name = rowC.GetProperty("Name").ToString() });
                    rowC1.SetFocus();
                    Playback.Wait(1000);
                    Globals.TempRow = rowC1;
                    return rowC1;
                }
            }
            return null;
        }

        public static string GetCellValueFromRow(WinRow rowC1, string cellHeader)
        {
            string cellValue = null;
            //var select = rowC1.Container.SearchFor<WinCell>(new { Name = cellHeader });
            //        select.SetFocus();
            //cellValue = select.GetProperty("Value").ToString();

            var cells = rowC1.Cells;

            foreach (var cell in cells)
            {
                if (cell.Name.Equals(cellHeader))
                {
                    cellValue = cell.GetProperty("Value").ToString();
                    return cellValue;
                }

            }
            return "";
        }

        public static bool SelectRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        {
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue != columnValue) continue;
                rowHeader.SetFocus();
                //Mouse.Click(rowHeader);
                var rowData = rowC.GetProperty("Name");
                var dataRow = table.Container.SearchFor<WinRow>(new { Name = rowData });
                dataRow.SetFocus();
                var rowHeader1 = dataRow.Container.SearchFor<WinCell>(new { Name = "Select" });
                //rowHeader.SetFocus();
                rowHeader1.SetFocus();
                Mouse.Click(rowHeader1);
                return true;
            }
            return false;
        }

        public static string GetStringFromSentence(string start, string end, string sentence)
        {
            return sentence.Substring((sentence.IndexOf(start) + start.Length), (sentence.IndexOf(end) - sentence.IndexOf(start) - start.Length));
        }


        public static void ClickButton(UITestControl windowProperties, string control)
        {
            var ctrl = Actions.GetWindowChild(windowProperties, control);
            Actions.MouseClickOnCoordinates(ctrl);
            Playback.Wait(1000);
            Actions.SendEnter();
        }

        public static void SelectRadioButton(UITestControl windowInst, string radioButtonGroupName, string radioButtonName)
        {
            var btnCont = Actions.GetWindowChild(windowInst, radioButtonGroupName);
            btnCont.SetFocus();
            var radioBtn = btnCont.Container.SearchFor<WinRadioButton>(new { Name = radioButtonName });
            radioBtn.SetFocus();
            Mouse.Click(radioBtn);

        }

        public static void SetMaskedText(UITestControl windowinst, string controlName, string data)
        {
            var control = Actions.GetWindowChild(windowinst, controlName);
            control.SetFocus();
            Playback.Wait(200);
            SendKeys.SendWait("{END}");
            Playback.Wait(200);
            SendKeys.SendWait("+{HOME}");
            Playback.Wait(200);
            SendKeys.SendWait("{DEL}");
            Playback.Wait(200);
            SendKeys.SendWait("{HOME}");
            Playback.Wait(200);
            SendKeys.SendWait(data);
        }
    }



    public static class Globals
    {
        public static string CustomerName { get; set; }
        public static string CustomerLegalName { get; set; }
        public static string CustomerContact { get; set; }
        public static string WorkerName { get; set; }
        public static string SSN { get; set; }
        public static string QuotedBy { get; set; }
        public static string JobOrderNo { get; set; }

        public static string Temp { get; set; }
        public static WinRow TempRow { get; set; }
        public static string Amount { get; set; }
    }

    public static class GlobalWindows
    {
        public static WinWindow CustomerProfileWindow { get; set; }
        public static WinWindow CustomerContactInfoWindow { get; set; }
        public static WinWindow CustomerSearchResultsWindow { get; set; }
        public static WinWindow CustomerWorkerComp { get; set; }
        public static WinWindow CustomerAddNote { get; set; }
        public static WinWindow COIPreviewWindow { get; set; }
        public static WinWindow DocumentOnFileWindow { get; set; }
        public static WinWindow DocumentOnFileFindFileWindow { get; set; }
        public static WinWindow QuoteProfileWindow { get; set; }
        public static WinWindow JobOrderProfileWindow { get; set; }

        public static WinWindow DialogWindow { get; set; }
    }

}