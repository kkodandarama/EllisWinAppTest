using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerAlreadyExistWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerAlreadyExistWindowProperties()
        {
            var existsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Worker Already Exist" });
            return existsWindow;
        }

        private static UITestControl GetTelephoneOverideWindowProperties()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            var tOveride = existsWindow.Container.SearchFor<WinWindow>(new { Name = "Telephone Number Override" });
            return tOveride;
        }

        #endregion

        #region Worker Already Exists Methods

        public static void ClickOnContinueBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var continueBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.ContinueBtn);
                MouseActions.ClickButton(existsWindow,
                    continueBtn.Enabled
                        ? WorkerexistsPopUpConstants.ContinueBtn
                        : WorkerexistsPopUpConstants.OverrideBtn);
            }
        }

        public static void ClickOnOverideBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var overrideBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.OverrideBtn);
                MouseActions.ClickButton(existsWindow,
                    overrideBtn.Enabled
                        ? WorkerexistsPopUpConstants.OverrideBtn
                        : WorkerexistsPopUpConstants.ContinueBtn);
            }
        }

        public static void ClickOnUpdateProfileBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var updateBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.UpdateProfileBtn);
                MouseActions.ClickButton(existsWindow,
                    updateBtn.Enabled
                        ? WorkerexistsPopUpConstants.UpdateProfileBtn
                        : WorkerexistsPopUpConstants.ContinueBtn);
            }
        }

        public static void ClickOnBackBtn()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
                MouseActions.ClickButton(existsWindow, WorkerexistsPopUpConstants.BackBtn);
        }

        public static void SelectWorkerFromGrid()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();
            if (existsWindow.Exists)
            {
                var row = TableActions.SelectRowFromTable(existsWindow, "grdWorkers", "WorkerInfoDomain row 1");
                Mouse.DoubleClick(row);
            }
        }

        public static void ClickonContinueBtnTelephone()
        {
            var tOveride = GetTelephoneOverideWindowProperties();
            if (tOveride.Exists)
                MouseActions.ClickButton(tOveride, "btnContinue");
        }

        public static bool VerifyUpdateBtnEnabled()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();

            var updateBtn = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.UpdateProfileBtn);
            return updateBtn.Enabled;
        }

        public static bool VerifyOverRideBtnEnabled()
        {
            var existsWindow = GetWorkerAlreadyExistWindowProperties();

            var overRide = Actions.GetWindowChild(existsWindow, WorkerexistsPopUpConstants.OverrideBtn);
            return overRide.Enabled;
        }

        #endregion

        #region Controls

        private class WorkerexistsPopUpConstants
        {
            public const string WorkersGrid = "grdWorkers";
            public const string ContinueBtn = "btnContinue";
            public const string OverrideBtn = "btnOverride";
            public const string UpdateProfileBtn = "btnUpdateProfile";
            public const string BackBtn = "btnBack";
        }

        #endregion
    }
}