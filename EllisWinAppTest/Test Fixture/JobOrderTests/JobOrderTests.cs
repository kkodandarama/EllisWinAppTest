using System;
using System.Diagnostics;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace EllisWinAppTest.JobOrderTests
{
    [CodedUITest]
    public class JoborderTests : AppContext
    {
        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CreateJobOrder()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);

            foreach (var dataRow in datarows.Where(x => x.ItemArray[1].ToString().Equals("CreateJobOrder")))
            {
                //Console.WriteLine(dataRow.ItemArray[1]);
                var jobOrderCreated = JobOrderWindow.CreateNewJobOrder(dataRow);
                Factory.AssertIsTrue(jobOrderCreated, "Job order not saved successfully");
                JobOrderWindow.CloseJobOrderProfileWindow();
            }
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void OpenExsistingJobOrder()
        {
            EllisHome.Initialize();
            var status = OpenJobOrder.SelectJobOrderFromTable();

            if (status)
                OpenJobOrder.CloseJobOrderProfile();

            Factory.AssertIsTrue(status, "Profile not found");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void OpenSpecificJobOrder()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.JobOrderVerify);
            foreach (var dataRow in dataRows.Where(x => !x.ItemArray[2].ToString().Equals(string.Empty)))
            {
                LandingPage.SelectFromToolbar("Job Orders");
                var status = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());

                if (status)
                    OpenJobOrder.CloseJobOrderProfile();

                Factory.AssertIsTrue(status, "Job Order with # " + dataRow.ItemArray[2].ToString() + " not found");
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CopyJobOrderDetails()
        {
            try
            {
                EllisHome.Initialize();

                LandingPage.SelectFromToolbar("Job Orders");
                CopyJobOrder.OpenAnyJobOrder();

                //Copy Job Order Details from opened job order
                var status = CopyJobOrder.CopyJobOrderDetails();
                if (status)
                    Factory.ClickButton(JobOrderWindow.GetNewJobOrderWindowProperties(), "btnCancel");
                OpenJobOrder.CloseJobOrderProfile();
                Factory.AssertIsTrue(status, "Job Order not copied successfully");
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CopyJobOrderAdditionalCharges()
        {
            try
            {
                EllisHome.Initialize();

                LandingPage.SelectFromToolbar("Job Orders");
                CopyJobOrder.OpenAnyJobOrder();

                //Copy Job Order Details from opened job order
                var status = CopyJobOrder.CopyJobOrderAdditionalCharges();
                if (status)
                    Factory.ClickButton(JobOrderWindow.GetNewJobOrderWindowProperties(), "btnCancel");
                OpenJobOrder.CloseJobOrderProfile();
                Factory.AssertIsTrue(status, "Job Order not copied successfully");
            }
            finally
            {
                Cleanup();
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CopyAndCreateJobOrder()
        {

            try
            {
                //Create job order from a copied details
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);
                foreach (var dataRow in datarows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("CopyJobOrder")))
                {
                    LandingPage.SelectFromToolbar("Job Orders");
                    //TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());
                    CopyJobOrder.OpenAnyJobOrder();
                    //Copy Job Order Details from opened job order
                    var status = CopyJobOrder.CopyJobOrderDetails();

                    if (status)
                    {
                        //Console.WriteLine(dataRow.ItemArray[1]);
                        //JobOrderWindow.EnterJobOrderData(dataRow);
                        //JobOrderWindow.ClickOnButton("Search");

                        //Playback.Wait(3000);
                        JobOrderWindow.ClickOnContinueBtn();
                        Windows.CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                        //// Find Quote Tab/Window
                        //Playback.Wait(3000);
                        //JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                        //JobOrderFindQuoteWindow.ClickOnButton("GO");
                        Playback.Wait(2000);
                        JobOrderWindow.ClickOnContinueBtn();

                        // Enter Basic Job Order Details
                        BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                        BasicJobInformationWindow.ClickOnContinueBtn();

                        status = PreQualifyingQuestionsWindow.HandleAlertWindow();
                        if (status)
                        {

                            Factory.AssertIsFalse(status, "Job Order alredy exist for this customer");
                        }
                        else
                        {
                            // Enter Schedule And Additional Charges Details
                            ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                            ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                            // Enter Order Notes in Schedule And Additional Charges window
                            ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                            // Focus back to Schedule And Additional Charges window
                            ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                            //Enter data in Requirements window
                            RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                            RequirementsWindow.ClickOnButton("Continue >");
                            Playback.Wait(3000);

                            //Enter data in Pre-Qualifying Requirements Window

                            PreQualifyingQuestionsWindow.ClickonSaveButton();
                            PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                            PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                            Playback.Wait(3000);
                            status = PreQualifyingQuestionsWindow.HandleAlertWindow();
                            Factory.AssertIsTrue(status, "Job order not saved successfully");
                            JobOrderWindow.CloseJobOrderProfileWindow();
                            JobOrderWindow.CloseJobOrderProfileWindow();
                        }

                    }
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CancelExistingJobOrder()
        {
            try
            {
                var dataRows = EllisHome.Initialize(ExcelFileNames.JobOrder);

                foreach (var data in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("CancelJobOrder")))
                {
                    if (data.ItemArray[77].ToString() != "" && data.ItemArray[78].ToString() != "")
                    {
                        LandingPage.SelectFromToolbar("Job Orders");
                        var recordStatus = CopyJobOrder.OpenAnyJobOrder();

                        if (recordStatus)
                        {
                            var joprofile = OpenJobOrder.JobOrderProfileWindowProperties();
                            MouseActions.ClickButton(joprofile, "btnCancelJobOrder");
                            CancelJobOrder.CancelNewJobOrder();
                            CancelJobOrder.EnterJobOrderNotes(data.ItemArray[77].ToString(), data.ItemArray[78].ToString());
                            var cancelStatus = CancelJobOrder.HandleAlertWindow();
                            Factory.AssertIsTrue(cancelStatus, "Job Order not canceled");

                            //Closing the newly created job order window
                            JobOrderWindow.CloseJobOrderProfileWindow();
                        }
                        else
                        {
                            Console.WriteLine("No Job order found.");
                        }
                    }
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CancelNewJobOrder()
        {
            try
            {
                var runStatus = string.Empty;
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);
                foreach (var dataRow in datarows.Where(x => x.ItemArray[1].Equals("CreateJobOrder")))
                {
                    //Data in "CancelJobOrderNotes" field is mandetory in TestData
                    if (dataRow.ItemArray[77].ToString() != String.Empty 
                        && dataRow.ItemArray[78].ToString() != String.Empty)
                    {
                        var jobOrderCreated = JobOrderWindow.CreateNewJobOrder(dataRow);
                        Factory.AssertIsTrue(jobOrderCreated, "Job order not saved successfully");
                        //Get job Order Number
                        Playback.Wait(3000);
                        Globals.JobOrderNo = JobOrderWindow.GetJobOrderNumber();
                        JobOrderWindow.CloseJobOrderProfileWindow();
                        //Cancel newly created job order
                        LandingPage.SelectFromToolbar("Job Orders");
                        TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", Globals.JobOrderNo);
                        var joprofile = OpenJobOrder.JobOrderProfileWindowProperties();
                        if (joprofile.Exists)
                        {
                            MouseActions.ClickButton(joprofile, "btnCancelJobOrder");
                            //CancelJobOrder.CancelNewJobOrder();
                            CancelJobOrder.EnterJobOrderNotes(dataRow.ItemArray[77].ToString(), dataRow.ItemArray[78].ToString());
                            var cancelStatus = CancelJobOrder.HandleAlertWindow();
                            Factory.AssertIsTrue(cancelStatus, "Job Order not canceled");

                            //Closing the newly created job order window
                            JobOrderWindow.CloseJobOrderProfileWindow();
                        }
                    }
                }
            }
            finally
            {
                Cleanup();
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void OpenExpiredJobOrder()
        {
            try
            {
                EllisHome.Initialize();
                LandingPage.SelectFromToolbar("Job Orders");
                var expJoProfile = false;
                //var status = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", dataRow.ItemArray[2].ToString());
                var status = OpenJobOrder.SelectExpiredJobOrderFromTable();
                if (status)
                {
                    //---------------------------------------------------------------------------------------------
                    //JoExpiration   
                    //---------------------------------------------------------------------------------------------
                    var lblControl = Actions.GetWindowChild(OpenJobOrder.JobOrderProfileWindowProperties(), OpenJobOrder.JobOrderSummaryConstents.JoExpiration);
                    Console.WriteLine("Job Order Expiry Date: " + Convert.ToDateTime(lblControl.GetProperty("Name").ToString()));
                    if (Convert.ToDateTime(lblControl.GetProperty("Name").ToString()) < DateTime.Now)
                    {
                        Console.WriteLine("Job Order Expired On: " + Convert.ToDateTime(lblControl.GetProperty("Name").ToString()));
                        expJoProfile = true;
                    }

                    Factory.AssertIsTrue(expJoProfile, "Expired Job Order Profile not found");
                    OpenJobOrder.CloseJobOrderProfile();
                }

                Factory.AssertIsTrue(status, "Profile not found");
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void EditExpiredJobOrder()
        {
            try
            {
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderEdit);
                var status = OpenJobOrder.SelectExpiredJobOrderFromTable();
                if (status)
                {
                    OpenJobOrder.CloseJobOrderProfile();
                }

                Factory.AssertIsTrue(status, "Profile not found");
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void VerifyJobOrderSummaryData()
        {
            try
            {
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderVerify);
                foreach (var data in datarows)
                {
                    LandingPage.SelectFromToolbar("Job Orders");

                    var profileStatus = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #",
                            data.ItemArray[2].ToString());
                    var status = true;
                    if (profileStatus)
                    {
                        status = OpenJobOrder.VerifyJobOrder(data);
                        //Factory.AssertIsTrue(status, "Profile data not matched");
                        OpenJobOrder.CloseJobOrderProfile();
                    }
                    Factory.AssertIsTrue(status, "Profile data not matched");
                    Factory.AssertIsTrue(profileStatus, "Profile not found");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void EditJobOrder()
        {
            try
            {
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrderEdit);
                foreach (var data in datarows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("EditJobOrder")))
                {
                    SearchWindow.SelectSearchElements(data.ItemArray[2].ToString(), "JobOrder",
                        SearchWindow.SearchTypeConstants.Simple);

                    //LandingPage.SelectFromToolbar("Job Orders");
                    var profileWindow = JobOrderWindow.GetNewJobOrderWindowProperties();
                    //var profileStatus = TableActions.OpenRecordFromTable(EllisWindow, "_grdJobOrders", "Job Order #", data.ItemArray[2].ToString());
                    if (profileWindow.Exists)
                    {
                        OpenJobOrder.SelectTab("Basic Job Info");
                        Playback.Wait(2000);
                        OpenJobOrder.EditBasicJobInfoOfJobOrder(data);

                        OpenJobOrder.SelectTab("OrderDetails/Addl Charges");
                        Playback.Wait(2000);
                        OpenJobOrder.EditOrderDetailsAddlChargesOfJobOrder(data);

                        OpenJobOrder.SelectTab("Requirements");
                        Playback.Wait(2000);
                        OpenJobOrder.EditRequirementsOfJobOrder(data);

                        OpenJobOrder.SelectTab("Pre-Qualifying Questions");
                        Playback.Wait(2000);
                        OpenJobOrder.EditPreQualifyingQuestionsOfJobOrder(data);

                        OpenJobOrder.SelectTab("Safety");
                        Playback.Wait(2000);
                        OpenJobOrder.EditSafetyOfJobOrder(data);

                        OpenJobOrder.SelectTab("Progress Billing");
                        Playback.Wait(2000);
                        OpenJobOrder.EditProgressBillingOfJobOrder(data);

                        OpenJobOrder.CloseJobOrderProfile();

                    }

                    Factory.AssertIsTrue(profileWindow.Exists, "Profile not found");

                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("JobOrder"), TestCategory("Positive")]
        public void CopyFromCreateJobOrder()
        {
            try
            {
                var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);

                foreach (
                    var dataRow in datarows.TakeWhile(dataRow => dataRow.ItemArray[1].ToString().Equals("CreateJobOrder")))
                {
                    if (dataRow.ItemArray[1].ToString().Equals("CreateJobOrder"))
                    {
                        //Console.WriteLine(dataRow.ItemArray[1]);
                        var jobOrderCreated = JobOrderWindow.CreateNewJobOrder(dataRow);
                        Factory.AssertIsTrue(jobOrderCreated, "Job order not saved successfully");

                        //Copy Job Order Details from opened job order
                        var status = CopyJobOrder.CopyJobOrderDetails();

                        if (status)
                            OpenJobOrder.CloseJobOrderProfile();
                        JobOrderWindow.ClickOnButton("Cancel");

                        Factory.AssertIsTrue(status, "Job Order not copied successfully");
                    }
                }
            }
            finally
            {
                Cleanup();
            }
        }

        private static void Cleanup()
        {
            EllisHome.App.Close();
        }
    }
}