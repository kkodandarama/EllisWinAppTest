using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class PreQualifyingQuestionsWindow : AppContext
    {
        private static UITestControl PreQualifyingQuestionsWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            //var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            return joborderWindow;
        }

        private static UITestControl VertexGeoCode()
        {
            var vertexGeoCode = App.Container.SearchFor<WinWindow>(new { ControlId = "CommonDialogView" });
            return vertexGeoCode;
        }

        private static UITestControl ChooseLocationWindowProperties()
        {
            var joborderWindow = App.Container.SearchFor<WinWindow>(new { Name = "JobOrder" });
            return joborderWindow;
        }

        public static UITestControl ChooseAlertWindowProperties()
        {
            var mainWindow = PreQualifyingQuestionsWindowProperties();
            var alertWindow = mainWindow.Container.SearchFor<WinWindow>(new { Name = "Alert" });
            return alertWindow;
        }

        public static void ClickonButton(string btnName)
        {
            Factory.ClickOnButton(PreQualifyingQuestionsWindowProperties(), btnName);
        }

        public static void HandleChooseLocationWindow()
        {
            var chooseLocationWindow = ChooseLocationWindowProperties();
            if (chooseLocationWindow.Exists)
            {
                //Factory.ClickOnButton(chooseLocationWindow, "Ok");
                MouseActions.ClickButton(chooseLocationWindow, "btnOK");
            }
        }

        public static void HandleWorkLocationWindow()
        {
            var chooseLocationWindow = ChooseLocationWindowProperties();

            if (chooseLocationWindow.Exists)
            {
                try
                {
                    var wininst = VertexGeoCode();
                    MouseActions.ClickButton(wininst, "_OKButton");
                }
                catch (Exception)
                {

                    MouseActions.ClickButton(chooseLocationWindow, "btnOK");
                }
            }
            
        }

        public static bool HandleAlertWindow()
        {
            var alertWindow = ChooseAlertWindowProperties();

            if (alertWindow.Exists)
            {
                //Factory.ClickOnButton(chooseAlertWindow, "OK");
                var okBtn = Actions.GetWindowChild(alertWindow, "_OKButton");
                Mouse.Click(okBtn);
                return true;
            }
            return false;
        }

        internal static void ClickonSaveButton()
        {
            Playback.Wait(3000);
            var preQuaWindow = PreQualifyingQuestionsWindowProperties();
            MouseActions.ClickButton(preQuaWindow, "FinishButton");

        }
    }
}