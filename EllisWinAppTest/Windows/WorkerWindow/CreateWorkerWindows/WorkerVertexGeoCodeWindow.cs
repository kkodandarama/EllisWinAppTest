using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerVertexGeoCodeWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerVertexGeoCodeWindowProperties()
        {
            var geoCodeWindow = App.Container.SearchFor<WinWindow>(new { Name = "Geo Code" });
            var vertexGeoCodeWindow = geoCodeWindow.Container.SearchFor<WinWindow>(new {ControlId = "CommonDialogView"});//{ Name = "New Applicant" });//CommonDialogView
            return vertexGeoCodeWindow;
        }

        #endregion

        #region Vertex Geo Code Methods

        public static void ClickOnOkBtn()
        {
            var vertexGeoCodeWindow = GetWorkerVertexGeoCodeWindowProperties();
            if (vertexGeoCodeWindow.Exists)
                MouseActions.ClickButton(vertexGeoCodeWindow, VertexGeoCodeConstants.OkBtn);
        }

        public static bool VerifyWorkerVertexGeoCodeWindowDisplayed()
        {
            var vertexGeoCodeWindow = GetWorkerVertexGeoCodeWindowProperties();
            return vertexGeoCodeWindow.Exists;
        }

        #endregion

        #region controls

        private class VertexGeoCodeConstants
        {
            public const string OkBtn = "_OKButton";
        }

        #endregion
    }
}