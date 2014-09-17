using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerCompleteBehavioralSurveryWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetWorkerSurveyWindowProperties()
        {
            var surveyWindow =
                App.Container.SearchFor<WinWindow>(new {ControlId = "BehavioralSurveyView"});    //{ ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return surveyWindow;
        }

        //public static UITestControl GetPopUpWindowProperties()
        //{
        //    WinWindow popWindow = new WinWindow();
        //    popWindow.TechnologyName = "MSAA";
        //    popWindow.SearchProperties[WinWindow.PropertyNames.ControlName] = "";
        //    //popWindow.SearchProperties[WinWindow.PropertyNames.ControlType] = "dialog";
        //    //popWindow.GetProperty("text").Equals("No record found for current SSN and current date-time.");

        //    return popWindow;
        //    //WinText innerText = new WinText(popWindow);
            
        //    ////string text = innerText.DisplayText;
        //    //if (innerText.DisplayText.Contains("No record found for current SSN and current date-time")) ;//Equals())
        //    //{
        //    //    return popWindow;
        //    //}
        //    //return null;
        //}
      
        //public static void ClickOnOkBtn()

        //{
        //    var popWin = GetPopUpWindowProperties();
        //    if (popWin.Exists)
        //    {
        //       Actions.SendAltF4();
        //    }
        //}

        //public static bool VerifyPopWindowExists()
        //{
        //    var popWin = GetPopUpWindowProperties();
        //    if (popWin.Exists)

        //    {
        //        return true;
        //    }
        //    return false;
        //}

        #endregion

        #region Complete Bahvioral Survey Methods

        public static void ClickOnGetResultsBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, SurveyConstants.GetResultsBtn);
        }

        public static void ClickOnBackBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, SurveyConstants.BackBtn);
        }

        public static void ClickOnCancelBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, SurveyConstants.CancelBtn);
        }

        public static void ClickOnRejectBtn()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            if (surveyWindow.Exists)
                MouseActions.ClickButton(surveyWindow, SurveyConstants.RejectBtn);
        }

        public static bool VerifyCompleteBehavioralSurveyWindowDisplayed()
        {
            var surveyWindow = GetWorkerSurveyWindowProperties();
            return surveyWindow.Exists;
        }

        #endregion

        #region Controls

        private class SurveyConstants
        {
            public const string RejectBtn = "btnReject";
            public const string GetResultsBtn = "btnGetResults";
            public const string CancelBtn = "btnCancel";
            public const string BackBtn = "_btnBack";
        }

        #endregion
    }
}