using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;
using System;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    public class JobOrderWindow : AppContext
    {
        public static void ClickOnCreateJobOrder()
        {
            var file = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.File });
            var joborder = file.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.JobOrder });
            var cjoborder = joborder.Container.SearchFor<WinMenuItem>(new { Name = EllisHomeConstants.CJobOrder });

            MouseActions.Click(file);
            MouseActions.Click(joborder);
            MouseActions.Click(cjoborder);
        }

        public static UITestControl GetNewJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return joborderWindow;
        }

        private static UITestControl GetCreateJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            return joborderWindow;
        }

        public static UITestControl GetJobOrderProfileProperties()
        {
            var window = App.Container.SearchFor<WinWindow>(new { Name = Globals.JobOrderNo + "-Job Order Profile" });
            return window;
        }

        public static UITestControl GetJobOrderDispatchPayoutWindowProperties(string name)
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = name });
            return joborderWindow;
        }

        public static bool VerifyCreateJobOrderWindowDisplayed()
        {
            var joWindow = GetCreateJobOrderWindowProperties();
            return joWindow.Exists;
        }

        public static bool VerifyDispatchStatusDisplayed()
        {
            var window = GetJobOrderProfileProperties();
            var label = window.Container.SearchFor<WinText>(new { ControlName = JobOrderConstants.JobOrderStatus });
            return label.Exists;
        }

        public static bool VerifyDispatchPayoutWindowDisplayed(string name)
        {
            var window = GetJobOrderDispatchPayoutWindowProperties(name);
            if (window.Exists)
                return true;
            return false;
        }

        private static UITestControlCollection GetJobOrderEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetCreateJobOrderWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
            var editControl = group.Container.SearchFor<WinEdit>(new { Name = "" });
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }


        public static void EnterJobOrderData(DataRow data)
        {
            var windowinst = GetCreateJobOrderWindowProperties();
            if (VerifyCreateJobOrderWindowDisplayed())
            {
                if (data.ItemArray[3].ToString() != "")
                    Actions.SetText(windowinst, JobOrderConstants.CustNumber, data.ItemArray[3].ToString());
                if (data.ItemArray[4].ToString() != "")
                    Actions.SetText(windowinst, JobOrderConstants.CustName, data.ItemArray[4].ToString());
                if (data.ItemArray[5].ToString() != "")
                    Factory.SetMaskedText(windowinst, JobOrderConstants.PhoneNumber, data.ItemArray[5].ToString());
                if (data.ItemArray[6].ToString() != "")
                    Factory.SetMaskedText(windowinst, JobOrderConstants.FebTaxId, data.ItemArray[6].ToString());
            }
            
        }

        public static void EnterJobOrderFindQuoteData(DataRow data)
        {
            var getJobOrderControlCollection = GetJobOrderEditControlCollection();

            foreach (var control in getJobOrderControlCollection)
            {
                if (control.FriendlyName != null)
                {
                    switch (control.FriendlyName)
                    {
                        case "Text area":
                            Actions.SetText(control, data.ItemArray[9].ToString());
                            break;
                        //case "mskFederalTaxId":
                        //    Actions.SendText("");
                        //    Actions.SendText("{HOME}");
                        //    //Actions.SendText(control, data.ItemArray[6].ToString());
                        //    Actions.SetText(control, data.ItemArray[6].ToString());
                        //    break;

                        default:
                            Console.WriteLine(control.FriendlyName);
                            break;
                    }
                }
            }
        }

        public static void ClickOnButton(string btnName)
        {
            MouseActions.ClickButton(GetCreateJobOrderWindowProperties(), "btnCancel");
            //Factory.ClickOnButton(GetCreateJobOrderWindowProperties(), btnName);

        }

        public static void ClickOnContinueBtn()
        {
            Playback.Wait(3000);
            var jobSearchWindow = GetCreateJobOrderWindowProperties();
            var continueBtn = Actions.GetWindowChild(jobSearchWindow, "btnCreateJobOrder");
            Mouse.Click(continueBtn);
        }

        //public static void ClickOnGoBtn()
        //{
        //    Playback.Wait(3000);
        //    var jobSearchWindow = GetCreateJobOrderWindowProperties();
        //    var butGroup = jobSearchWindow.Container.SearchFor<WinGroup>(new { Name = "" });
        //    var searchBtn = butGroup.Container.SearchFor<WinButton>(new { Name = "GO" });
        //    MouseActions.Click(searchBtn);
        //}

        public static void CloseJobOrderProfileWindow()
        {
            var newJobOrder = GetNewJobOrderWindowProperties();
            var joNum = Actions.GetWindowChild(newJobOrder, "lblJobOrderNumber");
            Console.WriteLine("Job Order #" + joNum.GetProperty("Name"));
            var closeBtn = Actions.GetWindowChild(newJobOrder, "btnClose");
            Mouse.Click(closeBtn);
        }

        public static bool CreateNewJobOrder(DataRow dataRow)
        {
            ClickOnCreateJobOrder();
            if (VerifyCreateJobOrderWindowDisplayed())
            {
                EnterJobOrderData(dataRow);
                MouseActions.ClickButton(OpenJobOrder.JobOrderProfileWindowProperties(), "btnSearch");
                //ClickOnButton("Search");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                
                ClickOnContinueBtn();
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                // Find Quote Tab/Window
                JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                ClickOnContinueBtn();

                // Enter Basic Job Order Details
                BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                BasicJobInformationWindow.ClickOnContinueBtn();

                // Enter Schedule And Additional Charges Details
                ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                // Enter Order Notes in Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                // Focus back to Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                //Enter data in Requirements window
                RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                //RequirementsWindow.ClickOnButton("Continue >");
                //Playback.Wait(3000);

                //Enter data in Pre-Qualifying Requirements Window
                //Playback.Wait(2000);
                //CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                PreQualifyingQuestionsWindow.ClickonSaveButton();
                Playback.Wait(2000);
                //CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                Playback.Wait(2000);
                //CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                //Playback.Wait(2000);
                PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                Playback.Wait(2000);
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                Playback.Wait(3000);

                var status = PreQualifyingQuestionsWindow.HandleAlertWindow();
                
                return status;
            }
            return false;
        }

        public static string GetJobOrderNumber()
        {
            var winInst = GetCreateJobOrderWindowProperties();
            var control = Actions.GetWindowChild(winInst, "lblJobOrderNumber");
            var jobOrderNo = control.GetProperty("Name").ToString();
            return jobOrderNo;
        }
    }

    public class JobOrderConstants
    {
        public const string CustNumber = "mskCustomerNumber";
        public const string CustName = "txtCustomerName";
        public const string PhoneNumber = "mskPhoneNumber";
        public const string FebTaxId = "mskFederalTaxId";
        public const string JobOrderStatus = "lblJobOrderStatus";
        // public const string CustName = "txtCustomerName";
        //public const string CustName = "txtCustomerName";
    }
}