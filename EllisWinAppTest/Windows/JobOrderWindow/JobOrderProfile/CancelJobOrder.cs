using System.Data;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile
{
    internal class CancelJobOrder : AppContext
    {
        public static UITestControl GetOrderNotesWindowProperties()
        {
            var mainWindow = JobOrderWindow.GetNewJobOrderWindowProperties();
            var ordrNotesWindow = mainWindow.Container.SearchFor<WinWindow>(new { Name = "Order Notes" });

            return ordrNotesWindow;
        }

        public static UITestControl GetAlertWindowProperties()
        {
            var mainWindow = JobOrderWindow.GetNewJobOrderWindowProperties();
            var ordrNotesWindow = mainWindow.Container.SearchFor<WinWindow>(new { Name = "Alert" });
            return ordrNotesWindow;
        }


        public static void EnterJobOrderNotes(string notes, string requestedBy)
        {
            var jobOrderWindow = GetOrderNotesWindowProperties();
            var sta = jobOrderWindow.Exists;
            var textOrderNotes = Actions.GetWindowChild(jobOrderWindow, "_txtOrderNotes");
            Actions.SetText(textOrderNotes, notes);

            var dropDownOrderNotes = Actions.GetWindowChild(jobOrderWindow, "_cmbRequestedBy");
            Playback.Wait(1000);
            DropDownActions.SelectDropdownByText(dropDownOrderNotes, requestedBy);

            var okBtn = Actions.GetWindowChild(jobOrderWindow, "_btnOk");
            Mouse.Click(okBtn);
        }

        public static bool HandleAlertWindow()
        {
            var alertWindow = GetAlertWindowProperties();

            if (alertWindow.Exists)
            {
                var okBtn = Actions.GetWindowChild(alertWindow, "_OKButton");
                Mouse.Click(okBtn);
                return true;
            }
            else
            {
                return false;
            }
        }

        public static void CancelNewJobOrder()
        {
            var newJobOrder = JobOrderWindow.GetNewJobOrderWindowProperties();
            MouseActions.ClickButton(newJobOrder, "btnCancelJobOrder");
        }
    }
}