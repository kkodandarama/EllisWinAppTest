using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest
{
    [CodedUITest]
    public class EllisWinTest : AppContext
    {
        [TestMethod]
        public void EllisClickOnFileExitTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllis();
            EllisHome.ClickOnFileExit();
            App.Close();
        }

        [TestMethod]
        public void LaunchEllisTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            Cleanup();
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
            App.Close();
        }
    }
}