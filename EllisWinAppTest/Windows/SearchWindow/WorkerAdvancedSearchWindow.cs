using System.Data;
using System.Linq;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class WorkerAdvancedSearchWindow : AppContext
    {


        #region Window Properties

        private static UITestControl GetWorkerSearchWindowProperties()
        {
            var workerSearchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            return workerSearchWindow;
        }

        public static UITestControl GetWorkerSearchResultsWindowProperties()
        {
            var workerSearchResultsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return workerSearchResultsWindow;
        }

        private static UITestControl GetWorkerApplicationOnFileWindowProperties()
        {
            var applicationonfile =
                App.Container.SearchFor<WinWindow>(new { Name = "WorkerApplicationOnFile" });
            return applicationonfile;
        }

        #endregion

        #region Worker Advanced Search Methods

        public static void SendTabs(int noOfTabs)
        {
            for (int tab = 0; tab < noOfTabs; tab++)
            {
                SendKeys.SendWait("{TAB}");
                Playback.Wait(1000);
            }
        }

        public static bool VerifyWorkerSearchWindowDisplayed()
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            return workerSearchWindow.Enabled;
        }

        public static bool EnterWorkerAdvancedSearchData(DataRow data)
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                var firstname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.FirstName);
                Actions.SetText(firstname, data.ItemArray[3].ToString());

                var lastname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.LastName);
                Actions.SetText(lastname, data.ItemArray[4].ToString());

                var address = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Address);
                Actions.SetText(address, data.ItemArray[5].ToString());

                var city = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.City);
                Actions.SetText(city, data.ItemArray[6].ToString());

                var state = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.State);
                if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                    DropDownActions.SelectDropdownByText(state, data.ItemArray[7].ToString());

                var zip = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.ZipCode);
                Actions.SetText(zip, data.ItemArray[8].ToString());

                var ssn = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.SSN);
                Actions.SetText(ssn, data.ItemArray[9].ToString());

                var workerNo = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.WorkerNumber);
                Actions.SetText(workerNo, data.ItemArray[10].ToString());

                var vehicle = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Vehicle);
                if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
                    DropDownActions.SelectDropdownByText(vehicle, data.ItemArray[14].ToString());

                var wotc = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Wotc);
                if (!string.IsNullOrEmpty(data.ItemArray[15].ToString()))
                    DropDownActions.SelectDropdownByText(wotc, data.ItemArray[15].ToString());

                var primaryStatus = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.PrimaryStatus);
                if (!string.IsNullOrEmpty(data.ItemArray[16].ToString()))
                    DropDownActions.SelectDropdownByText(primaryStatus, data.ItemArray[16].ToString());

                var reason = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Reason);
                if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
                    DropDownActions.SelectDropdownByText(reason, data.ItemArray[17].ToString());

                var workDesc = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.WokrDesc);
                Actions.SetText(workDesc, data.ItemArray[18].ToString());

                var branch = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Branch);
                if (!string.IsNullOrEmpty(data.ItemArray[19].ToString()))
                    DropDownActions.SelectDropdownByText(branch, data.ItemArray[19].ToString());

                var lastWork = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.LatsWorkDt);
                if (!string.IsNullOrEmpty(data.ItemArray[20].ToString()))
                    DropDownActions.SelectDropdownByText(lastWork, data.ItemArray[20].ToString());

                var dispatchFrom = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.DispatchFrom);
                if (!string.IsNullOrEmpty(data.ItemArray[21].ToString()))
                    DropDownActions.SelectDropdownByText(dispatchFrom, data.ItemArray[21].ToString());

                var dispatchTo = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.DispatchTo);
                if (!string.IsNullOrEmpty(data.ItemArray[22].ToString()))
                    DropDownActions.SelectDropdownByText(dispatchTo, data.ItemArray[22].ToString());

                var applicantFrom = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.ApplicantFrom);
                if (!string.IsNullOrEmpty(data.ItemArray[23].ToString()))
                    DropDownActions.SelectDropdownByText(applicantFrom, data.ItemArray[23].ToString());

                var applicantTo = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.ApplicantTo);
                if (!string.IsNullOrEmpty(data.ItemArray[24].ToString()))
                    DropDownActions.SelectDropdownByText(applicantTo, data.ItemArray[24].ToString());

                var ticketNo = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.TicketNo);
                Actions.SetText(ticketNo, data.ItemArray[25].ToString());

                var joAddress = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.JobAddress);
                Actions.SetText(joAddress, data.ItemArray[26].ToString());

                var jobCity = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.JobCity);
                Actions.SetText(jobCity, data.ItemArray[27].ToString());

                var jobState = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.JobState);
                if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                    DropDownActions.SelectDropdownByText(jobState, data.ItemArray[28].ToString());

                var jobZip = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.JobZip);
                Actions.SetText(jobZip, data.ItemArray[29].ToString());

                var i9From = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.I9Start);
                if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                    DropDownActions.SelectDropdownByText(i9From, data.ItemArray[30].ToString());

                var i9To = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.I9End);
                if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                    DropDownActions.SelectDropdownByText(i9To, data.ItemArray[31].ToString());

                var eVerify = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.Everify);
                if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                    DropDownActions.SelectDropdownByText(eVerify, data.ItemArray[32].ToString());

                return true;
            }

            return false;

        }

        public static void EnterWorkerNameAsSearchData(string name)
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();

            var firstname = Actions.GetWindowChild(workerSearchWindow, WorkerSearchConstants.FirstName);
            Actions.SetText(firstname, name);

            MouseActions.ClickButton(workerSearchWindow, WorkerSearchConstants.SearchBtn);
        }

        public static bool ClickCancelBtn()
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchWindow, WorkerSearchConstants.CancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickSearchBtn()
        {
            var workerSearchWindow = GetWorkerSearchWindowProperties();
            if (workerSearchWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchWindow, WorkerSearchConstants.SearchBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Worker Search Results Methods

        public static bool ClickOnRefineSearchBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.RefineSearchBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnPrintBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.PrintBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnExportBtn()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                MouseActions.ClickButton(workerSearchResultsWindow, WorkerSearchResultsConstatnts.ExportBtn);
                return true;
            }
            return false;
        }

        public static bool CloseSearchResultsWindow()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)workerSearchResultsWindow);
                return true;
            }
            return false;
        }

        public static bool SelectWorkerfromSearchResults(string workernumber)
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            if (workerSearchResultsWindow.Exists)
            {
                var label = Actions.GetWindowChild(workerSearchResultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var workerNo = TableActions.OpenRecordFromTable(workerSearchResultsWindow, WorkerSearchResultsConstatnts.SearchGrid,
                    "Worker Number", workernumber);
                return workerNo;
            }
            return false;
        }

        public static bool VerifyWorkerSearchResultsWindowDisplayed()
        {
            var workerSearchResultsWindow = GetWorkerSearchResultsWindowProperties();
            return workerSearchResultsWindow.Enabled;
        }

        #endregion

        #region Worker Application On File methods

        public static void SelectApplicationChkBox(DataRow data)
        {
            var applicationonfile = GetWorkerApplicationOnFileWindowProperties();
            if (applicationonfile.Exists)
            {
                var checkBox = Actions.GetWindowChild(applicationonfile, "chkApplicationReceived");
                if (!string.IsNullOrEmpty(data.ItemArray[35].ToString()))
                {
                    Actions.SetCheckBox((WinCheckBox)checkBox, data.ItemArray[35].ToString());
                    MouseActions.ClickButton(applicationonfile, "btnOK");
                }
            }
        }

        #endregion

        #region Controls

        private class WorkerSearchConstants
        {
            public const string LastName = "txtLastName";
            public const string ZipCode = "txtWorkerZipCode";
            public const string City = "txtWorkerCity";
            public const string Address = "txtWorkerAddress";
            public const string FirstName = "txtFirstName";
            public const string WorkerNumber = "txtWorkerNumber";
            public const string SSN = "txtSSN";
            public const string State = "cmbWorkerState";
            public const string PositionFocus = "lstPositionFocus";
            public const string Title = "lstTitle";
            public const string Skills = "lstSkillsExperience";
            public const string Vehicle = "cmbVehicle";
            public const string Wotc = "cmbWOTCStatus";
            public const string PrimaryStatus = "cmbWorkerPrimaryStatus";
            public const string Reason = "cmbWorkerSecondaryStatus";
            public const string WokrDesc = "txtWorkDescription";
            public const string Branch = "cmbBranch";
            public const string LatsWorkDt = "dteLastWorkeddate";
            public const string DispatchFrom = "dteFirstDispatchdate";
            public const string DispatchTo = "dteFirstDispatchToDate";
            public const string ApplicantFrom = "dteDate";
            public const string ApplicantTo = "dteEnteredApplicantToDate";
            public const string TicketNo = "txtTicketNumber";
            public const string JobAddress = "txtJobAddressLine1";
            public const string JobCity = "txtJobSiteCity";
            public const string JobState = "cmbJobSiteState";
            public const string JobZip = "txtJobSiteZipCode";
            public const string I9Start = "i9ExpirationStart";
            public const string I9End = "i9ExpirationEnd";
            public const string Everify = "cmbEVerifyStatus";
            public const string SearchBtn = "_buttonSearch";
            public const string CancelBtn = "_cancelButton";
        }

        private class WorkerSearchResultsConstatnts
        {
            public const string SearchGrid = "_grdSearchResult";
            public const string RefineSearchBtn = "_btnRefineSearch";
            public const string PrintBtn = "_btnPrint";
            public const string ExportBtn = "_btnExport";

        }

        #endregion
    }
}