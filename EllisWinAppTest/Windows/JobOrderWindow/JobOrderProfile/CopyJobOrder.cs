using System;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.CustomerTests;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile
{
    internal class CopyJobOrder : AppContext
    {
        public static bool CopyJobOrderDetails()
        {
            var jobOrderWindow = OpenJobOrder.JobOrderProfileWindowProperties();

            if (jobOrderWindow.Exists)
            {
                MouseActions.ClickButton(jobOrderWindow, "btnCopyJobOrder");
                //Factory.ClickOnButton(jobOrderWindow, "Copy Job Order");
                var copyWindow = GetCopyOptionsWindowProperties();

                var chkBox = Actions.GetWindowChild(copyWindow, "_chkCopyJobOrderDetails");
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");
                MouseActions.ClickButton(copyWindow, "_btnOk");
                //Factory.ClickOnButton(copyWindow, "OK");
                return true;
            }
            return false;

        }

        public static bool CopyJobOrderAdditionalCharges()
        {
           var jobOrderWindow = OpenJobOrder.JobOrderProfileWindowProperties();
            
           if (jobOrderWindow.Exists)
            {
                Factory.ClickButton(jobOrderWindow, "btnCopyJobOrder");
                var copyWindow = GetCopyOptionsWindowProperties();

                var chkBox = Actions.GetWindowChild(copyWindow, "_chkCopyAdditionalCharges");
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");

                //Factory.ClickOnButton(copyWindow, "OK");
                Factory.ClickButton(copyWindow, "_btnOk");
                return true;
            }
            return false;
        }

        public static UITestControl GetCopyOptionsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            return winGroup;
        }

        public static bool OpenAnyJobOrder()
        {
            var calRange = Actions.GetWindowChild(EllisWindow, "btnToggle");
            if (calRange.GetProperty("Name").Equals("Advanced..."))
                Mouse.Click(calRange);
            var dateOnly = DateTime.Now.GetDateTimeFormats();
            var dateTo = dateOnly[3];
            var date1 = DateTime.Now.AddDays(-45).GetDateTimeFormats();
            var date1from = date1[3];

            Factory.SetMaskedText(EllisWindow, "advancedFromDate", date1from.Replace("/", ""));
            Factory.SetMaskedText(EllisWindow, "advancedToDate", dateTo.Replace("/", ""));
            SendKeys.SendWait("{TAB}");
            Playback.Wait(2000);

            var status = Factory.OpenFirstRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #");
            //Handle Warning popup
            return status;

        }
    }
}