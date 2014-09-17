using System;
using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerConfirmApplicantElgibiltyWindow :AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerConfirmationWindowProperties()
        {
            var cWorkerWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "WorkerConfirmationView"});//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });//
            return cWorkerWindow;
        }

        #endregion

        #region Confirmation Methods

        public static void ClickOnYesBtn()
        {
            var cWorkerWindow = GetWorkerConfirmationWindowProperties();
            if (cWorkerWindow.Exists)
                MouseActions.ClickButton(cWorkerWindow, ConfirmationConstants.YesBtn);
        }

        public static void ClickOnNoBtn()
        {
            var cWorkerWindow = GetWorkerConfirmationWindowProperties();
            if (cWorkerWindow.Exists)
                MouseActions.ClickButton(cWorkerWindow, ConfirmationConstants.NoBtn);
        }

        public static bool VerifyConfirmationPopUpDisplayed()
        {
            var cWorkerWindow = GetWorkerConfirmationWindowProperties();
            if (cWorkerWindow.Exists)
            {
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        private class ConfirmationConstants
        {
            public const string YesBtn = "btnYes";
            public const string NoBtn = "btnNo";
        }

        #endregion
    }
}
