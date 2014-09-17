using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.DispatchAndPayoutWindow
{
    class OpenByStatus : AppContext
    {
        public static void OpenDispatchAndPayoutWindow(string dispatchStatus)
        {
            EllisWindow.WaitForControlReady();
            var view = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var dispOrPout = view.Container.SearchFor<WinMenuItem>(new { Name = "Dispatch / Payout" });
            var window = dispOrPout.Container.SearchFor<WinMenuItem>(new { Name = dispatchStatus });

            MouseActions.Click(view);
            MouseActions.Click(dispOrPout);
            MouseActions.Click(window);
            
        }        
    }

}
