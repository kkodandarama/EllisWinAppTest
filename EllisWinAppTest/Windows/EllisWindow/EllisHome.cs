using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml.Linq;

namespace EllisWinAppTest.Windows.EllisWindow
{
    public class EllisHome : AppContext
    {
        public static IEnumerable<DataRow> Initialize(string excelName, int retry = 5)
        {
            WindowsActions.KillEllisProcesses();
            App = LaunchEllisAsCSRUser();
            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return Initialize(excelName, retry - 1);
            var datarows = ExcelReader.ImportSpreadsheet(excelName);
            return datarows;
        }

        public static void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = LaunchEllisAsCSRUser();
            App.Process.WaitForInputIdle(6000);
        }

        public static ApplicationUnderTest LaunchEllis()
        {
            App = ApplicationUnderTest.Launch(TestData.Path, TestData.AltPath);
            Factory.GetWinProperties();

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsCSRUser(int retry = 5)
        {
            //LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            LaunchAppAsDifferentUser(TestData.Path, TestData.CSRUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();
            var ready = App.WaitForControlExist(5000);

            if (!ready && retry >= 0)
                return LaunchEllisAsCSRUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsARMUser(int retry = 5)
        {
            //LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            LaunchAppAsDifferentUser(TestData.Path, TestData.ARMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);

            if (!ready && retry >= 0)
                return LaunchEllisAsCSRUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsDMUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.DMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsDMUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsARRUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.ARRUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsARRUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsAVPUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.AVPUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsAVPUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsNABSUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.NABSUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsNABSUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsNAPSUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.NAPSUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsNAPSUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsNAAMUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.NAAMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsNAAMUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsPRMUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.PRMUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsPRMUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsTAUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.TAUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsTAUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsRAUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.RAUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsRAUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsPWUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.PWUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsPWUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsGSUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.GSUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsGSUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsGLUser(int retry = 5)
        {
            LaunchAppAsDifferentUser(TestData.Path, TestData.GLUsername, TestData.AppPassword,
                TestData.AppDomain);
            App = Factory.GetAUT();
            Factory.GetWinProperties();

            var ready = App.WaitForControlExist(5000);
            if (!ready && retry >= 0)
                return LaunchEllisAsGLUser(retry - 1);

            return App;
        }

        public static ApplicationUnderTest LaunchEllisAsDiffUserFromDesktop()
        {
            LaunchActions.LaunchAppAsDiffferentUserFromDesktop();
            Playback.Wait(6000);
            App = Factory.GetAUT();
            Factory.GetWinProperties();
            return App;
        }

        public static void ClickOnFileExit()
        {
            try
            {
                var ellisWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Ellis"));
                var closeButton = ellisWindow.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Close"));
                Mouse.Click(closeButton.GetClickablePoint());
            }
            catch (Exception)
            {
                //Suppress any exception here
            }
        }

        public static void ClickOnFilePayrollOCR()
        {
            EllisWindow.SetFocus();
            var menu = EllisWindow.Container.SearchFor<WinMenuBar>(new { Name = "Toolbar" });
            var file = menu.Items[0];
            Mouse.Click(file);
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Payroll });
            Mouse.Click(payroll);
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Garnishments });
            Mouse.Click(garnishment);
            var ocr = garnishment.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.OCR });
            Mouse.Click(ocr);
        }

        public static void ClickOnFilePayrollCreateOrder()
        {
            EllisWindow.SetFocus();
            var menu = EllisWindow.Container.SearchFor<WinMenuBar>(new { Name = "Toolbar" });
            var file = menu.Items[0];
            Mouse.Click(file);
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Payroll });
            Mouse.Click(payroll);
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Garnishments });
            Mouse.Click(garnishment);
            var order = garnishment.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Order });
            Mouse.Click(order);
        }

        public static void ClickOnVieworPrintAnswerLetters()
        {
            EllisWindow.SetFocus();
            var menu = EllisWindow.Container.SearchFor<WinMenuBar>(new { Name = "Toolbar" });
            var file = menu.Items[0];
            Mouse.Click(file);
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Payroll });
            Mouse.Click(payroll);
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Garnishments });
            Mouse.Click(garnishment);
            var letter = garnishment.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.AnswerLetter });
            Mouse.Click(letter);
        }

        public static void ClickOnRateSheet()
        {
            EllisWindow.SetFocus();
            var menu = EllisWindow.Container.SearchFor<WinMenuBar>(new { Name = "Toolbar" });
            var file = menu.Items[0];
            Mouse.Click(file);
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Payroll });
            Mouse.Click(payroll);
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.RateSheet });
            Mouse.Click(garnishment);

        }

        public static void ClickOnOrders()
        {
            EllisWindow.SetFocus();
            //var menu = EllisWindow.Container.SearchFor<WinMenuBar>(new { Name = "Toolbar" });
            var views = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });

            var payroll = views.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });

            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Garnishments });

            var orders = garnishment.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Orders });

            Mouse.Click(views);
            Mouse.Click(payroll);
            Mouse.Click(garnishment);
            Mouse.Click(orders);

        }

        public static void OpenPrevailingJobWindow(string prevailingWageJobStatus)
        {
            var view = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var payRoll = view.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var prevailingWage = payRoll.Container.SearchFor<WinMenuItem>(new { Name = "Prevailing Wage" });
            var window = prevailingWage.Container.SearchFor<WinMenuItem>(new { Name = prevailingWageJobStatus });

            MouseActions.Click(view);
            MouseActions.Click(payRoll);
            MouseActions.Click(prevailingWage);
            MouseActions.Click(window);

        }

        public static void ClickOnUnClaimedProperty()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "File" });
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var unclaimedProperty = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Unclaimed Property" });

            Mouse.Click(file);
            Mouse.Click(payroll);
            Mouse.Click(unclaimedProperty);
        }

        public static void ClickOnOvertime()
        {
            var views = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var payroll = views.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var overTime = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Overtime" });

            Mouse.Click(views);
            Mouse.Click(payroll);
            Mouse.Click(overTime);
        }

        public static void ClickOnStateHoliday()
        {
            var views = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var payroll = views.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var stateHoliday = payroll.Container.SearchFor<WinMenuItem>(new { Name = "State Holiday" });

            Mouse.Click(views);
            Mouse.Click(payroll);
            Mouse.Click(stateHoliday);
        }

        public static void ClickOnPayment()
        {
            var views = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var payroll = views.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.Garnishments });
            var payment = garnishment.Container.SearchFor<WinMenuItem>(new { Name = "Payments" });

            Mouse.Click(views);
            Mouse.Click(payroll);
            Mouse.Click(garnishment);
            Mouse.Click(payment);
        }

        public static void ClickOnPayrollAdjustments()
        {
            var views = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });

            var payroll = views.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });

            var pAdjust = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Payroll Adjustments" });

            var payrollAdjust = pAdjust.Container.SearchFor<WinMenuItem>(new { Name = "Payroll Adjustments" });

            Mouse.Click(views);
            Mouse.Click(payroll);
            Mouse.Click(pAdjust);
            SendKeys.SendWait("{UP}");
            SendKeys.SendWait("{ENTER}");
            //Mouse.Click(payrollAdjust);
        }

        public static void ClickOnSearchPayments()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "File" });
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var searchPayments = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Search Payments" });

            Mouse.Click(file);
            Mouse.Click(payroll);
            Mouse.Click(searchPayments);
        }

        public static void ClickOnClearedChecksGateway()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "File" });
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var gateWay = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Cleared Checks Gateway" });

            Mouse.Click(file);
            Mouse.Click(payroll);
            Mouse.Click(gateWay);
        }

        public static void ClickOnSchedulePayments()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "File" });
            var payroll = file.Container.SearchFor<WinMenuItem>(new { Name = "Payroll" });
            var garnishment = payroll.Container.SearchFor<WinMenuItem>(new { Name = "Garnishments" });
            var schedule = garnishment.Container.SearchFor<WinMenuItem>(new { Name = "Schedule Payments" });

            Mouse.Click(file);
            Mouse.Click(payroll);
            Mouse.Click(garnishment);
            Mouse.Click(schedule);
        }

        private static ApplicationUnderTest LaunchAppAsDifferentUser(string appPath, string username, string password,
            string domain)
        {
            var isBranch = IsThisSessionABranchSession(appPath);

            var s = new NetworkCredential("", password).SecurePassword;
            var path = @appPath;
            var myProc = new ProcessStartInfo(path);

            try
            {
                myProc.UseShellExecute = false;
                myProc.WorkingDirectory = Path.GetDirectoryName(path);
                if (!isBranch)
                {
                    myProc.Domain = domain;
                    myProc.UserName = username;
                    myProc.Password = s;
                    myProc.LoadUserProfile = true;
                }
            }
            catch (Exception)
            {
                // error handling
            }

            App = ApplicationUnderTest.Launch(myProc);

            if (isBranch)
            {
                App.WaitForControlReady(6000);
                var logOnScreenHandled = HandleLaunchScreen(domain, username, password);
                Factory.AssertIsTrue(logOnScreenHandled, "The Log On Screen handler failed.");
            }

            return App;
        }

        /// <summary>
        /// This checks the AppSettings.config file for an "add" Element with a "IsBranch" Attribute.  It looks for the "value" value.
        /// </summary>
        /// <param name="appPath">The full filepath to the Ellis.exe</param>
        /// <returns>The IsBranch value (True or False).  False if the Attribute isn't in that file.  False if that file doesn't exist.</returns>
        private static bool IsThisSessionABranchSession(string appPath)
        {
            var fullPath = Path.Combine(
                Path.GetDirectoryName(appPath),
                "Config",
                "ClientConfig",
                "AppSettings.config");

            bool isBranch = false;
            if (File.Exists(fullPath))
            {
                try
                {
                    var settings = XElement.Load(fullPath);
                    var isBranchXElement = settings.Elements("add").Where(s => s.Attribute("key").Value == "IsBranch").FirstOrDefault();
                    if (isBranchXElement != null)
                        isBranch = Boolean.Parse(isBranchXElement.Attribute("value").Value);
                }
                catch (Exception e)
                {
                    Debug.WriteLine("Couldn't read the AppSettings.config file. \n" + e);
                    return false;
                }
            }

            return isBranch;
        }

        private static bool HandleLaunchScreen(string domain, string userName, string password)
        {
            try
            {
                var launcherScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "EllisLauncherScreen"));
                if (launcherScreen == null) return true;
                var differenUserRadioButton = launcherScreen.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "radioButtonDifferentUser"));
                var userNameTextEdit = launcherScreen.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "comboBoxUserName"));
                var passwordTextEdit = launcherScreen.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "textBoxPassword"));
                var logonButton = launcherScreen.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "buttonLogon"));
                var exitButton = launcherScreen.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "buttonExit"));

                (differenUserRadioButton.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern).Select();
                (userNameTextEdit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern).SetValue(String.Format("{0}\\{1}", domain, userName));
                (passwordTextEdit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern).SetValue(password);
                (logonButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern).Invoke();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                return false;
            }

            return true;
        }
    }

    internal class EllisHomeConstants
    {
        public const string File = "File";
        public const string Exit = "Exit";
        public const string Payroll = "Payroll";
        public const string Worker = "Workers";
        public const string CApplicant = "Create Applicant";
        public const string JobOrder = "Job Orders";
        public const string CJobOrder = "Create Job Order";

        public const string Customer = "Customers";
        public const string CCustomer = "Create Customer";

        public const string Garnishments = "Garnishments";
        public const string OCR = "Import OCR Batch";
        public const string Order = "Create New Order";
        public const string AnswerLetter = "View/Print Answer Letters";
        public const string RateSheet = "Rate Sheet";
        public const string Orders = "Orders";
    }
}