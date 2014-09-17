using System;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.WorkerWindow.CreateWorkerWindows
{
    public class WorkerGeoCodeWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetWorkerGeoCodeWindowProperties()
        {
            var geoCodeWindow = App.Container.SearchFor<WinWindow>(new {ControlId = "VertexGeoCodeView"});//{Name = "Geo Code"});
            return geoCodeWindow;
        }

        #endregion

        #region Geo Code Methods

        public static void ClickOnOkBtn()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            if (geoCodeWindow.Exists)
                MouseActions.ClickButton(geoCodeWindow, GeoCodeConstants.OkBtn);
        }

        public static bool VerifyGeoCodeWindowDisplayed()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            return geoCodeWindow.Exists;
        }

        public static void ClickOnCancelBtn()
        {
            var geoCodeWindow = GetWorkerGeoCodeWindowProperties();
            if (geoCodeWindow.Exists)
                MouseActions.ClickButton(geoCodeWindow, GeoCodeConstants.CancelBtn);
        }

        #endregion

        #region Controls

        private class GeoCodeConstants
        {
            public const string OkBtn = "btnOK";
            public const string CancelBtn = "btnCancel";
            public const string GeoCodeGrid = "grdGeoCode";
            public const string GeoCodeRow = "VertexGeoCodeDomain row 1";
        }

      
        #endregion
    }
}