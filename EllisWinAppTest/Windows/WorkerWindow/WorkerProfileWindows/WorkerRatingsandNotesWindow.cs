using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerRatingsandNotesWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerProfileWindowProperties()
        {
            var workerProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            return workerProfileWindow;
        }

        private static UITestControl GetTempWorkerNotesWindowProperties()
        {
            var notesWindow = App.Container.SearchFor<WinWindow>(new { Name = "Temporary Worker Notes" });
            return notesWindow;
        }

        private static UITestControl GetTempWorkerRatingsWindowProperties()
        {
            var ratingsWindow = App.Container.SearchFor<WinWindow>(new { Name = "Temporary Worker Ratings" });
            return ratingsWindow;
        }

        private static UITestControl GetCustomerSearchWindowProperties()
        {
            var cSearchWindow = App.Container.SearchFor<WinWindow>(new { Name = "Customer Search" });
            return cSearchWindow;
        }

        #endregion

        #region Ratings and Notes Methods

        public static void SelectRatingsInComboBox()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var ratingsCmbbox = Actions.GetWindowChild(workerProfileWindow, WorkerRatingsandNotesTabConstants.Ratings);
            DropDownActions.SelectDropdownByText(ratingsCmbbox, "Ratings");

            MouseActions.ClickButton(workerProfileWindow, WorkerRatingsandNotesTabConstants.AddRatingsBtn);

        }

        public static void SelectNotesInComboBox()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();

            var ratingsCmbbox = Actions.GetWindowChild(workerProfileWindow, WorkerRatingsandNotesTabConstants.Ratings);
            MouseActions.Click(ratingsCmbbox);
            SendKeys.SendWait("Notes");
            Playback.Wait(1000);

            MouseActions.ClickButton(workerProfileWindow, WorkerRatingsandNotesTabConstants.AddNotesBtn);
        }

        public static bool ClickAddRatingsBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var addRatingsBtn = Actions.GetWindowChild(workerProfileWindow,
                WorkerRatingsandNotesTabConstants.AddRatingsBtn);
                MouseActions.Click(addRatingsBtn);
                return true;
            }
            return false;
        }

        public static bool ClickAddNotesBtn()
        {
            var workerProfileWindow = GetWorkerProfileWindowProperties();
            if (workerProfileWindow.Exists)
            {
                var addNotesBtn = Actions.GetWindowChild(workerProfileWindow,
                WorkerRatingsandNotesTabConstants.AddNotesBtn);
                MouseActions.Click(addNotesBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Ratings window Methods

        public static bool VerifyRatingsWindowDisplayed()
        {
            var ratingsWindow = GetTempWorkerRatingsWindowProperties();
            return ratingsWindow.Enabled;
        }

        public static void EnterdataInRatingsWindow(DataRow data)
        {
            var ratingsWindow = GetTempWorkerRatingsWindowProperties();

            var ratingType = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.RatingType);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                DropDownActions.SelectDropdownByText(ratingType, data.ItemArray[3].ToString());

            var jobOrder = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.JobOrder);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                DropDownActions.SelectDropdownByText(jobOrder, data.ItemArray[4].ToString());

            var comments = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.Comments);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                Actions.SetText(comments, data.ItemArray[5].ToString());

            MouseActions.ClickButton(ratingsWindow, TempWorkerRatingsWindowConstants.SubmitBtn);

        }

        public static bool ClickCancelRatings()
        {
            var ratingsWindow = GetTempWorkerRatingsWindowProperties();
            if (ratingsWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickSubmitRatings()
        {
            var ratingsWindow = GetTempWorkerRatingsWindowProperties();
            if (ratingsWindow.Exists)
            {
                var submitBtn = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.SubmitBtn);
                Mouse.Click(submitBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCustomerBrowseBtn()
        {
            var ratingsWindow = GetTempWorkerRatingsWindowProperties();
            if (ratingsWindow.Exists)
            {
                var browseBtn = Actions.GetWindowChild(ratingsWindow, TempWorkerRatingsWindowConstants.BrowseBtn);
                Mouse.Click(browseBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Notes Window Methods

        public static bool VerifyNotesWindowDisplayed()
        {
            var notesWindow = GetTempWorkerNotesWindowProperties();
            return notesWindow.Enabled;
        }

        public static void EnterdataInNotesWindow(DataRow data)
        {
            var notesWindow = GetTempWorkerNotesWindowProperties();

            var notesType = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.NoteType);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                DropDownActions.SelectDropdownByText(notesType, data.ItemArray[6].ToString());

            var jobOrder = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.JobOrder);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                DropDownActions.SelectDropdownByText(jobOrder, data.ItemArray[7].ToString());

            var comments = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.Comments);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                Actions.SetText(comments, data.ItemArray[8].ToString());

            MouseActions.ClickButton(notesWindow, TempWorkerNotesWindowConstants.SubmitBtn);

        }

        public static bool ClickCancelnotes()
        {
            var notesWindow = GetTempWorkerNotesWindowProperties();
            if (notesWindow.Exists)
            {
                var cancelBtn = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.CancelBtn);
                Mouse.Click(cancelBtn);
                return true;
            }
            return false;
        }

        public static bool ClickSubmitnotes()
        {
            var notesWindow = GetTempWorkerNotesWindowProperties();
            if (notesWindow.Exists)
            {
                var submitBtn = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.SubmitBtn);
                Mouse.Click(submitBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnBrowseBtn()
        {
            var notesWindow = GetTempWorkerNotesWindowProperties();
            if (notesWindow.Exists)
            {
                var browseBtn = Actions.GetWindowChild(notesWindow, TempWorkerNotesWindowConstants.BrowseBtn);
                Mouse.Click(browseBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Customer Search window Methods

        public static bool VerifyCustomerSearchWindowDisplayed()
        {
            var cSearchWindow = GetCustomerSearchWindowProperties();
            return cSearchWindow.Enabled;
        }

        public static void EnterDataCustomerSearchWindow(DataRow data)
        {
            var cSearchWindow = GetCustomerSearchWindowProperties();

            var customerNumber = Actions.GetWindowChild(cSearchWindow, CustomerSearchWindowConstants.CustomerNumber);
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            Actions.SetText(customerNumber, data.ItemArray[9].ToString());
            
            MouseActions.ClickButton(cSearchWindow, CustomerSearchWindowConstants.SearchCustomer);
            Playback.Wait(1000);

            MouseActions.ClickButton(cSearchWindow, CustomerSearchWindowConstants.CloseBtn);

        }

        public static bool ClickOnCloseBtn()
        {
            var cSearchWindow = GetCustomerSearchWindowProperties();
            if (cSearchWindow.Exists)
            {
                var closeBtn = Actions.GetWindowChild(cSearchWindow, CustomerSearchWindowConstants.CloseBtn);
                Mouse.Click(closeBtn);
                return true;
            }
            return false;
        }

        public static bool ClickOnCustomerNoBtn()
        {
            var cSearchWindow = GetCustomerSearchWindowProperties();
            if (cSearchWindow.Exists)
            {
                var browseBtn = Actions.GetWindowChild(cSearchWindow, CustomerSearchWindowConstants.SearchCustomer);
                Mouse.Click(browseBtn);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class WorkerRatingsandNotesTabConstants
        {
            public const string Ratings = "cmbRatings";
            public const string RatingsGrid = "grdRatings";
            public const string AddRatingsBtn = "btnRatings";
            public const string AddNotesBtn = "btnNotes";
        }

        private class TempWorkerNotesWindowConstants
        {
            public const string NoteType = "cmbCRatingType";
            public const string Customer = "txtCustomer";
            public const string JobOrder = "cmbJobOrder";
            public const string BrowseBtn = "btnBrowse";
            public const string CancelBtn = "btnCancel";
            public const string SubmitBtn = "btnSubmit";
            public const string Comments = "txtComments";
        }

        private class TempWorkerRatingsWindowConstants
        {
            public const string RatingType = "cmbRatingType";
            public const string Customer = "txtCustomer";
            public const string JobOrder = "cmbJobOrder";
            public const string BrowseBtn = "btnCustomer";
            public const string CancelBtn = "btnCancel";
            public const string SubmitBtn = "btnSubmit";
            public const string Comments = "txtComments";
        }

        private class CustomerSearchWindowConstants
        {
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string SearchCustomer = "btnSearchCustomer";
            public const string CloseBtn = "btnClose";
        }

        #endregion
    }
}