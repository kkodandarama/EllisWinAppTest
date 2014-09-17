using System;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.DispatchAndPayoutWindow
{
    class PrintDispatch : AppContext
    {
        public static UITestControl PrintOptionsWindowProperties()
        {
            var printWindow = App.Container.SearchFor<WinWindow>(new { Name = "Print Dispatch Sheet" });
            return printWindow;

        }

        public static string HandleEllisExceptionWindow()
        {
            var exceptionWindow = App.Container.SearchFor<WinWindow>(new { Name = "Ellis Exception Information" });
            if (exceptionWindow.Exists)
            {
                var exceptionTextControl = exceptionWindow.Container.SearchFor<WinText>(new { ControlName = "_messageLabel" });
                var exception = exceptionTextControl.GetProperty("Name").ToString();
                MouseActions.ClickButton(exceptionWindow, "ButtonAccept");
                return exception;
            }
            return null;
        }

        public static void HandleErrorWindow()
        {
            var exceptionWindow = App.Container.SearchFor<WinWindow>(new { Name = "ERROR" });
            if (exceptionWindow.Exists)
            {
                MouseActions.ClickButton(exceptionWindow, "_OKButton");
            }
        }

        public static bool PrintDispatchWithMap()
        {
            MouseActions.ClickButton(DispatchProfileWindow.DispatchProfileWindowProperties(), "btnPrintDispatchSheet");
            HandleErrorWindow();
            DispatchProfileWindow.HandleErrorMessageWindow();
            var printWindowControl = PrintOptionsWindowProperties();
            //Actions.SelectRadioButton(printWindowControl, "Need Map");
            Playback.Wait(5000);
            MouseActions.ClickButton(printWindowControl, "btnPrint");
            
            var exp = HandleEllisExceptionWindow();
            if (exp == null)
            {
                HandleErrorWindow();
                int i = 0;
                var print = printWindowControl.Container.SearchFor<WinWindow>(new { Name = "Print" });
                //var print1 = print.GetChildren();
                //print1[3].SetFocus();
                //Playback.Wait(3000);
                //Mouse.Click(print1[3]);
                //Playback.Wait(3000);
                while (print.Exists && i < 10)
                {
                    print.SetFocus();
                    SendKeys.SendWait("{ESC}");
                    Playback.Wait(2000);
                    i++;
                    print = App.Container.SearchFor<WinWindow>(new { Name = "Print" });
                }
                return true;

            }
            return false;
        }

        public static bool PrintDispatchWithOutMap()
        {
            MouseActions.ClickButton(DispatchProfileWindow.DispatchProfileWindowProperties(), "btnPrintDispatchSheet");

            var printWindowControl = PrintOptionsWindowProperties();
            //Actions.SelectRadioButton()
            //Actions.SelectRadioButton(printWindowControl, "Do Not Need Map");
            Playback.Wait(2000);
            SendKeys.SendWait(" ");

            MouseActions.ClickButton(printWindowControl, "btnPrint");
            Playback.Wait(5000);
            var exp = HandleEllisExceptionWindow();
            if (exp == null)
            {
                int i = 0;
                var print = App.Container.SearchFor<WinWindow>(new { Name = "Print" });
                while (print.Exists && i < 10)
                {
                    print.SetFocus();
                    SendKeys.SendWait("{ESC}");
                    Playback.Wait(2000);
                    i++;
                    print = App.Container.SearchFor<WinWindow>(new { Name = "Print" });
                }

                return true;
            }
            else
            {
                Console.WriteLine(exp);
                MouseActions.ClickButton(printWindowControl, "btnCancel");

                return false;
            }
        }
    }
}
