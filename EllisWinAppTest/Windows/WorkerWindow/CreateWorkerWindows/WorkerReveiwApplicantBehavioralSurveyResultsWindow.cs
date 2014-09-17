using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerReveiwApplicantBehavioralSurveyResultsWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerSurveyWindowProperties()
        {
            var surveyWindow =
                App.Container.SearchFor<WinWindow>(new { ControlId = "BehavioralSurveyQualifiedView" });//{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });//Name = "Survey-" + Globals.WorkerName });
            return surveyWindow;
        }

        private static UITestControl GetWorkerQualifiedSurveyWindowProperties()
        {
            var task = Task<WinWindow>.Factory.StartNew(
                () => App.Container.SearchFor<WinWindow>(
                    new { ControlId = "BehavioralSurveyPreviouslyQualifiedView" }));

            var completed = task.Wait(new System.TimeSpan(0, 0, 12));
            return completed ? task.Result : null;
        }

        private static UITestControl GetWorkerNonQualifiedWindowProperties()
        {
            var nonQualified = App.Container.SearchFor<WinWindow>(new { ControlId = "BehavioralSurveyNonQualifiedView" });
            return nonQualified;
        }

        private static UITestControl GetWorkerSurveyResultsWindowProperties()
        {
            var results = App.Container.SearchFor<WinWindow>(new { ControlId = "BehavioralSurveyResultsView" });
            return results;
        }

        private static UITestControl GetprintWindowProperties()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            var printWindow = surveyWindow.Container.SearchFor<WinWindow>(new { Name = "Print" });
            return printWindow;
        }

        #endregion

        #region Review Bahvioral Survey results Methods

        public static void ClickOnContinueBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.ContinueBtn);
        }

        public static bool VerifyContinueBtnDisplayed()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();

            var continueBtn = Actions.GetWindowChild(surveyWindow, ReviewConstants.ContinueBtn);
            if (continueBtn.Exists)
            {
                return true;
            }
            return false;
        }

        public static void ClickOnBackBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.BackBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.CancelBtn);
        }

        public static void ClickOnNextBtnQualified()
        {
            var qualified = GetWorkerQualifiedSurveyWindowProperties();
            if (qualified.IsControlUsable())
                MouseActions.ClickButton(qualified, "btnContinue");
            
        }

        #endregion

        #region Tescor Methods

        public static void ClickonContinueBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyResultsWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.ContinueBtntescor);
        }

        public static void ClickonCancelBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyResultsWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.CancelBtntescor);
        }

        public static void ClickonBackBtnTescor()
        {
            var surveyWindow = GetWorkerSurveyResultsWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.BackBtntescor);
        }

        public static void ClickonPrintBtn()
        {
            var surveyWindow = GetWorkerNonQualifiedWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.PrintLetterBtn);
        }

        public static void ClosePrintWindow()
        {
            var printWindow = GetprintWindowProperties();
            if (printWindow.Exists)
            {
                printWindow.SetFocus();
                SendKeys.SendWait("{ESC}");
            }
        }

        public static bool VerifyPrintWindowDisplayed()
        {
            var printWindow = GetprintWindowProperties();
            return printWindow.Exists;
        }

        public static void ClickonCloseBtn()
        {
            var surveyWindow = GetWorkerNonQualifiedWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, ReviewConstants.CloseBtn);
        }

        public static void EnterDatainTescor()
        {
            var surveyWindow = GetWorkerSurveyResultsWindowProperties();
            if (surveyWindow.Exists)
            {
                var tescorSsn = Actions.GetWindowChild(surveyWindow, ReviewConstants.TescorSsn);
                tescorSsn.SetFocus();
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{TAB}");

                var ssn = Actions.GetWindowChild(surveyWindow, ReviewConstants.Ssn);
                ssn.SetFocus();
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{TAB}");

                var chkBox = Actions.GetWindowChild(surveyWindow, ReviewConstants.DateFilterChkBx);
                Actions.SetCheckBox((WinCheckBox)chkBox, "TRUE");
            }
        }

        public static bool VerifyWorkerNonQualifiedWindowDisplayed()
        {
            var nonQualified = GetWorkerNonQualifiedWindowProperties();
            return nonQualified.Exists;
        }

        public static bool VerifyWorkerSurveyResultsWindowDisplayed()
        {
            var results = GetWorkerSurveyResultsWindowProperties();
            return results.Exists;
        }

        #endregion

        #region Controls

        private class ReviewConstants
        {
            public const string ContinueBtn = "btnContinue";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "_btnBack";
            public const string TescorSsn = "cmbTescorSSN";
            public const string Ssn = "cmbSSN";
            public const string SearchDate = "dtpSearchDate";
            public const string TryagainBtn = "_btnTryAgain";
            public const string DateFilterChkBx = "cbUseDateFilter";
            public const string BackBtntescor = "btnBack";
            public const string ContinueBtntescor = "btnContinue";
            public const string CancelBtntescor = "btnCancel";
            public const string PrintLetterBtn = "btnPrintLetter";
            public const string CloseBtn = "btnClose";

        }

        #endregion
    }
}
