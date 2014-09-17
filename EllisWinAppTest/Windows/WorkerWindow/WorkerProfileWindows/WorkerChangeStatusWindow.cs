using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.WorkerProfileWindows
{
    public class WorkerChangeStatusWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetChangeStatusWindowProperties()
        {
            //var workerProfileWindow =
            //    App.Container.SearchFor<WinWindow>(new { Name = "Worker Profile-" + Globals.WorkerName });
            var changeStatusWindow = App.Container.SearchFor<WinWindow>
                (new { Name = "Worker Profile" });
            return changeStatusWindow;
        }

        #endregion

        #region Status Window Methods

        public static bool VerifyChangeStatusWindowDisplayed()
        {
            var changeStatusWindow = GetChangeStatusWindowProperties();
            return (changeStatusWindow.Enabled);
        }

        public static bool ClickOnCancelBtnStatusWindow()
        {
            var changeStatusWindow = GetChangeStatusWindowProperties();
            if (changeStatusWindow.Exists)
            {
                MouseActions.ClickButton(changeStatusWindow, "btnCancel");
                return true;
            }
            return false;

        }

        public static bool ClickOnStatusDropDown()
        {
            var changeStatusWindow = GetChangeStatusWindowProperties();
            if (changeStatusWindow.Exists)
            {
                var status = Actions.GetWindowChild(changeStatusWindow, StatusWorkerConstants.PrimaryStatus);
                MouseActions.Click(status);
                Playback.Wait(2000);
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class StatusWorkerConstants
        {
            public const string PrimaryStatus = "cmbPrimaryStatus";
            public const string Reason = "cmbSecondaryStatus";
            public const string EffectiveDate = "dteInactiveDate";
        }

        #endregion
    }
}