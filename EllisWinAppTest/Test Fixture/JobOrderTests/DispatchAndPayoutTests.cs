using System;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Windows.Interop;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.DispatchAndPayoutWindow;
using Microsoft.VisualStudio.TestTools.Common;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace EllisWinAppTest.JobOrderTests
{
    [CodedUITest]
    public class DispatchAndPayoutTests : AppContext
    {
        # region Dispatch Profile


        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void DispatchWorker()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("DispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("Assign Workers");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");

                var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                DispatchProfileWindow.AddWorkerBySimpleSearch(dataRow);
                DispatchProfileWindow.HandleWorkerSearchResultsWindow(dataRow);

                # region Assign worker to a JobOrder

                // Select Worker from Grid: grdOrderDetails
                var workerName = dataRow.ItemArray[7] + ", " + dataRow.ItemArray[6];
                if (dataRow.ItemArray[8].ToString() != String.Empty)
                    workerName = workerName + "," + dataRow.ItemArray[8];

                var workerFound = TableActions.SelectCellFromTable(dispatchProfile, "grdOrderDetails", "Worker",
                    workerName);
                var setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                Actions.SetCheckBox((WinCheckBox)setWeek, "True");
                Playback.Wait(1000);

                var control = Actions.GetWindowChild(dispatchProfile, "btnAssignWorker");
                Mouse.Click(control);

                var messageWindow = DispatchProfileWindow.ValidationMessageWindowProperties();
                if (messageWindow.IsControlUsable())
                    MouseActions.ClickButton(messageWindow, "_OKButton");

                //Handle save windows
                var winInst = DispatchProfileWindow.AssignWorkerWindowProperties();
                if (winInst.IsControlUsable())
                    MouseActions.ClickButton(winInst, "btnOK");

                winInst = DispatchProfileWindow.AssignAckWindowProperties();
                if (winInst.IsControlUsable())
                    MouseActions.ClickButton(winInst, "_OKButton");
                # endregion

                # region Dispatch worker to a JobOrder
                // Select Worker from Grid: grdOrderDetails
                workerName = dataRow.ItemArray[7] + ", " + dataRow.ItemArray[6];
                if (dataRow.ItemArray[8].ToString() != String.Empty)
                    workerName = workerName + "," + dataRow.ItemArray[8];

                workerFound = TableActions.SelectCellFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);

                //Check the box
                setWeek = SelectWorkerFromGridCheckbox(dispatchProfile, setWeek, dataRow);

                MouseActions.ClickButton(dispatchProfile, "btnPrintWorkTicket");
                DispatchProfileWindow.HandleTicketNotPrinted();

                DispatchProfileWindow.HandlePrintWorkTicketWindow(dataRow);

                # endregion

                //# region Payout for the assigned worker

                DispatchProfileWindow.SelectTab("Payout");

                //Select Worker from Grid: grdOrderDetails
                DispatchProfileWindow.SetDataInPayoutTableCell(dataRow);
                DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                ClickButtonCalcSave();
                Playback.Wait(1200);
                MouseActions.ClickButton(DispatchProfileWindow.ValidationMessageWindowProperties(), "_OKButton");
                DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                MouseActions.ClickButton(dispatchProfile, "btnCalcPrint");
                MouseActions.ClickButton(DispatchProfileWindow.ValidationMessageWindowProperties(), "_OKButton");

                //Click Print Email Payout
                DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                MouseActions.ClickButton(dispatchProfile, "btnPrintPayout");
                DispatchProfileWindow.HandlePrintWorkerPayoutWindow(dataRow);

                // Review Dispatch 
                DispatchProfileWindow.SelectTab("Review Dispatch");
                TableActions.SelectCellFromTable(dispatchProfile, "grdDispatchReview", "Worker", workerName);
                setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                Actions.SetCheckBox((WinCheckBox)setWeek, "True");
                MouseActions.ClickButton(dispatchProfile, "btnMarkReviewed");
                MouseActions.ClickButton(DispatchProfileWindow.ReviewDispatchWindowProperties(), "_OKButton");
                DispatchProfileWindow.SelectTab("Dispatch");
                MouseActions.ClickButton(dispatchProfile, "btnCancel");
            }

        }

        private static bool ClickButtonCalcSave()
        {
            var windows = AutomationElement.RootElement.FindAll(
                TreeScope.Descendants, new PropertyCondition(
                    AutomationElement.IsWindowPatternAvailableProperty, true));

            AutomationElement window = AutomationElement.RootElement;

            foreach (AutomationElement w in windows)
            {
                if (w.Current.Name.EndsWith("Dispatch / Payout"))
                {
                    window = w;
                    break;
                }
            }

            var select = window.FindFirst(
                TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "btnCalcSave"));

            object pattern;

            if (select.TryGetCurrentPattern(InvokePattern.Pattern, out pattern))
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

        private static UITestControl SelectWorkerFromGridCheckbox(UITestControl dispatchProfile, UITestControl setWeek, System.Data.DataRow dataRow)
        {
            setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
            Actions.SetCheckBox((WinCheckBox)setWeek, "True");

            var tableName = Actions.GetWindowChild(DispatchProfileWindow.DispatchProfileWindowProperties(), "grdOrderDetails");
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = "Worker" });
                rowHeader.SetFocus();
                var callValue = rowHeader.GetProperty("Value").ToString();

                var workerName = dataRow.ItemArray[7] + ", " + dataRow.ItemArray[6];
                if (dataRow.ItemArray[8].ToString() != String.Empty)
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
                        return select;
                    }
                }
            }

            return null;
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void DispatchWorkerAndSendEmail()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("DispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("Assign Workers");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                EllisWinAppTest.Windows.CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                    dispatchProfile.SetFocus();
                    DispatchProfileWindow.AddWorkerBySimpleSearch(dataRow);
                    Playback.Wait(2000);
                    DispatchProfileWindow.HandleWorkerSearchResultsWindow(dataRow);

                    # region Assign worker to a JobOrder

                    // Select Worker from Grid: grdOrderDetails

                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker",
                        workerName);
                    if (workerFound)
                    {
                        var setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                        Actions.SetCheckBox((WinCheckBox)setWeek, "True");
                        Playback.Wait(1000);
                        // MouseActions.ClickButton(dispatchProfile, "btnAssignWorker");
                        var control = Actions.GetWindowChild(dispatchProfile, "btnAssignWorker");
                        control.SetFocus();
                        Mouse.Click(control);

                        MouseActions.ClickButton(DispatchProfileWindow.AssignWorkerWindowProperties(), "btnOK");
                        MouseActions.ClickButton(DispatchProfileWindow.AssignAckWindowProperties(), "_OKButton");

                    # endregion

                        # region Dispatch worker to a JobOrder

                        // Select Worker from Grid: grdOrderDetails
                        workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                        if (dataRow.ItemArray[8].ToString() != String.Empty)
                            workerName = workerName + "," + dataRow.ItemArray[8];

                        workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker",
                            workerName);

                        if (workerFound)
                        {
                            setWeek = Actions.GetWindowChild(dispatchProfile, "ChkWeek");
                            Actions.SetCheckBox((WinCheckBox)setWeek, "True");
                            // Send Email
                            MouseActions.ClickButton(dispatchProfile, "btnPrintWorkTicket");

                            DispatchProfileWindow.HandleEmailWorkTicketWindow(dataRow);

                            MouseActions.ClickButton(DispatchProfileWindow.AssignAckWindowProperties(), "_OKButton");

                        # endregion
                        }

                        //# region Payout for the assigned worker

                        DispatchProfileWindow.SelectTab("Payout");

                        //Select Worker from Grid: grdOrderDetails
                        DispatchProfileWindow.SetDataInPayoutTableCell(dataRow);
                        MouseActions.ClickButton(dispatchProfile, "btnCalcSave");
                        if (DispatchProfileWindow.SaveWindowProperties().Exists)
                            MouseActions.ClickButton(DispatchProfileWindow.SaveWindowProperties(), "OKBTN");
                        DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                        MouseActions.ClickButton(dispatchProfile, "btnCalcPrint");
                        MouseActions.ClickButton(DispatchProfileWindow.ValidationMessageWindowProperties(), "_OKButton");

                        MouseActions.ClickButton(dispatchProfile, "btnCancel");

                        MouseActions.ClickButton(dispatchProfile, "btnCancel");
                    }

                }
            }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void WorkerPayoutCalcSave()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("DispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);

                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();

                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = TableActions.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker",
                       workerName);
                    if (workerFound)
                    {

                        DispatchProfileWindow.SelectTab("Payout");

                        //Select Worker from Grid: 
                        DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                        //add customer signeture
                        var ddControl = Actions.GetWindowChild(dispatchProfile, "cmbSignature");
                        DropDownActions.SelectDropdownByText(ddControl, dataRow.ItemArray[38].ToString());

                        //add customer signeture notes
                        var notesControl = Actions.GetWindowChild(dispatchProfile, "txtNotes");
                        Actions.SetText(notesControl, dataRow.ItemArray[39].ToString());

                        MouseActions.ClickButton(dispatchProfile, "btnCalcSave");
                        Assert.IsTrue(DispatchProfileWindow.HandleSelectWorkerValidationWindow());

                        MouseActions.ClickButton(dispatchProfile, "btnCancel");
                    }
                }


            }

        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void WorkerPayoutCalcPrint()
        {
            try
            {
                var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
                var dataRow = dataRows.Where(x => x.ItemArray[1].ToString().Equals("DispatchWorker")).FirstOrDefault();
                Factory.AssertIsFalse(dataRow == null, "Couldn't find DispatchWorker in the data.");
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);

                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();

                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = TableActions.SelectRecordFromTable(
                        dispatchProfile,
                        "grdOrderDetails",
                        "Worker",
                        workerName);

                    if (workerFound)
                    {

                        DispatchProfileWindow.SelectTab("Payout");

                        //Select Worker from Grid: 
                        DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                        //add customer signeture
                        var ddControl = Actions.GetWindowChild(dispatchProfile, "cmbSignature");
                        DropDownActions.SelectDropdownByText(ddControl, dataRow.ItemArray[38].ToString());

                        //add customer signeture notes
                        var notesControl = Actions.GetWindowChild(dispatchProfile, "txtNotes");
                        Actions.SetText(notesControl, dataRow.ItemArray[39].ToString());

                        MouseActions.ClickButton(dispatchProfile, "btnCalcSave");
                        Assert.IsTrue(DispatchProfileWindow.HandleSelectWorkerValidationWindow());
                        DispatchProfileWindow.SelectPayoutTableColumns(dataRow);
                        MouseActions.ClickButton(dispatchProfile, "btnCalcPrint");
                        Assert.IsTrue(DispatchProfileWindow.HandleSelectWorkerValidationWindow());

                        MouseActions.ClickButton(dispatchProfile, "btnCancel");
                        if (dispatchProfile.Exists)
                            MouseActions.ClickButton(dispatchProfile, "btnCancel");
                    }
                }

            }
            finally
            {
                Cleanup();
            }

        }



        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void UnAssignWorker()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            bool noRecordsFound = true;
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("UnAssignWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");

                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);

                if (dispatchRecordFound)
                {
                    noRecordsFound = false;
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();

                    # region Un-Assign worker to a JobOrder

                    // Select Worker from Grid: grdOrderDetails
                    DispatchProfileWindow.UnAssignUnDispatchWorker(dataRow);
                    var removeStatus = DispatchProfileWindow.RemoveDispatchWorker(dataRow);
                    if (removeStatus)
                        Console.WriteLine("Worker successfully removed from Job Order");
                    Console.WriteLine("Worker not removed from Job Order, either worker is Assigned or Dispatch");

                    Factory.AssertIsTrue(removeStatus, "Worker not removed from Job Order, either worker is Assigned or Dispatch");
                    # endregion
                    MouseActions.ClickButton(dispatchProfile, "btnCancel");
                }
            }

            Assert.IsFalse(noRecordsFound, "There were no Dispatch Records found, so we didn't UnAssign a Worker.");
        }



        # endregion

        # region Print Dispatch Profile

        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void PrintDispatchWithMap()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("DispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                Factory.AssertIsTrue(dispatchRecordFound, "Couldn't find the Dispatch Record.");
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);
                    Factory.AssertIsTrue(workerFound, "Couldn't find the worker.");
                    if (workerFound)
                    {
                        DispatchProfileWindow.SelectDispatchTableColumns(dataRow);
                        Assert.IsTrue(PrintDispatch.PrintDispatchWithMap(), "Printing Dispatch Profile failed");

                        MouseActions.ClickButton(dispatchProfile, "btnCancel");
                    }
                }
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void PrintDispatchWithOutMap()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("DispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = Factory.SelectRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);
                    if (workerFound)
                    {
                        DispatchProfileWindow.SelectDispatchTableColumns(dataRow);
                        Assert.IsTrue(PrintDispatch.PrintDispatchWithOutMap(), "Printing Dispatch Profile failed");
                        MouseActions.ClickButton(dispatchProfile, "btnCancel");
                    }
                }
            }
        }

        #endregion

        # region Verify Dispatch
        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void VerifyDispatchWorkerBillRateDetails()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            foreach (var dataRow in dataRows.Where(dataRow => dataRow.ItemArray[1].ToString().Equals("VerifyDispatchWorker")))
            {
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                    dispatchProfile.SetFocus();


                    // Select Worker from Grid: grdOrderDetails

                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];

                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];
                    Globals.WorkerName = workerName;

                    var workerFound = TableActions.OpenRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);
                    if (workerFound)
                    {
                        # region Compare Bill Rate


                        var row = DispatchProfileWindow.GetRowFromTable("Bill Rate");

                        if (!dataRow.ItemArray[10].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[10].ToString().Replace("\n", String.Empty), dataRow.ItemArray[24].ToString());

                        if (!dataRow.ItemArray[12].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[12].ToString().Replace("\n", String.Empty), dataRow.ItemArray[25].ToString());

                        if (!dataRow.ItemArray[14].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[14].ToString().Replace("\n", String.Empty), dataRow.ItemArray[26].ToString());

                        if (!dataRow.ItemArray[16].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[16].ToString().Replace("\n", String.Empty), dataRow.ItemArray[27].ToString());


                        if (!dataRow.ItemArray[18].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[18].ToString().Replace("\n", String.Empty), dataRow.ItemArray[28].ToString());

                        if (!dataRow.ItemArray[20].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[20].ToString().Replace("\n", String.Empty), dataRow.ItemArray[29].ToString());

                        if (!dataRow.ItemArray[22].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[22].ToString().Replace("\n", String.Empty), dataRow.ItemArray[30].ToString());


                        # endregion
                    }
                }

                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
            }
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void VerifyDispatchWorkerPayRateDetails()
        {
            try 
	        {	        
		        var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
                var dataRow = dataRows.FirstOrDefault(x => x.ItemArray[1].ToString().Equals("VerifyDispatchWorker"));
                Factory.AssertIsFalse(dataRow == null, "Couldn't find DispatchWorker in the data.");
                OpenByStatus.OpenDispatchAndPayoutWindow("All");
                var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
                if (dispatchRecordFound)
                {
                    var dispatchProfile = DispatchProfileWindow.DispatchProfileWindowProperties();
                    //dispatchProfile.SetFocus();

                    // Select Worker from Grid: grdOrderDetails
                    var workerName = dataRow.ItemArray[7] + "," + dataRow.ItemArray[6];
                    Globals.WorkerName = workerName;
                    if (dataRow.ItemArray[8].ToString() != String.Empty)
                        workerName = workerName + "," + dataRow.ItemArray[8];

                    var workerFound = TableActions.OpenRecordFromTable(dispatchProfile, "grdOrderDetails", "Worker", workerName);
                    Factory.AssertIsTrue(workerFound, "Couldn't find the worker.");
                    if (workerFound)
                    {
                        var row = DispatchProfileWindow.GetRowFromTable("Pay Rate");

                        if (!dataRow.ItemArray[10].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[10].ToString().Replace("\n", String.Empty), dataRow.ItemArray[31].ToString());

                        if (!dataRow.ItemArray[12].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[12].ToString().Replace("\n", String.Empty), dataRow.ItemArray[32].ToString());

                        if (!dataRow.ItemArray[14].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[14].ToString().Replace("\n", String.Empty), dataRow.ItemArray[33].ToString());

                        if (!dataRow.ItemArray[16].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[16].ToString().Replace("\n", String.Empty), dataRow.ItemArray[34].ToString());


                        if (!dataRow.ItemArray[18].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[18].ToString().Replace("\n", String.Empty), dataRow.ItemArray[35].ToString());

                        if (!dataRow.ItemArray[20].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[20].ToString().Replace("\n", String.Empty), dataRow.ItemArray[36].ToString());

                        if (!dataRow.ItemArray[22].ToString().Equals(String.Empty))
                            DispatchProfileWindow.ComapreBillRate(row, dataRow.ItemArray[22].ToString().Replace("\n", String.Empty), dataRow.ItemArray[37].ToString());
                    }
                }
            
                Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
	        }
	        finally
	        {
                    Cleanup();
	        }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("DispatchandPayout"), TestCategory("Positive")]
        public void VerifyDispatchByStatus()
        {
            var dataRows = EllisHome.Initialize(ExcelFileNames.Dispatch);
            var dataRow = dataRows.Where(
                x => x.ItemArray[1].Equals("VerifyDispatch")
                && !string.IsNullOrWhiteSpace(x.ItemArray[5].ToString()))
                .FirstOrDefault();

            OpenByStatus.OpenDispatchAndPayoutWindow(dataRow.ItemArray[5].ToString());
            var dispatchRecordFound = DispatchProfileWindow.OpenJobOrderForDispatch(dataRow);
            Factory.AssertIsTrue(dispatchRecordFound, "JobOrder not found for dispatching worker");
            MouseActions.ClickButton(DispatchProfileWindow.DispatchProfileWindowProperties(), "btnCancel");
        }



        # endregion


        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }

    }
}
