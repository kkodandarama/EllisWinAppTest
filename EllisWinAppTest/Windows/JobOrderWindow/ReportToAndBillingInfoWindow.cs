using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class ReportToAndBillingInfoWindow : AppContext
    {
        private static UITestControl ReportToAndBillingInfoWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Create New JobOrder"});
            return joborderWindow;
        }

        private static UITestControlCollection GetReportToAndBillingInfoWindowEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ReportToAndBillingInfoWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetReportToAndBillingInfoWindowDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ReportToAndBillingInfoWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetReportToAndBillingInfoWindowRadioButtonControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = ReportToAndBillingInfoWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinRadioButton>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        public static UITestControlCollection GetButtonColloction(UITestControl windowInstence, string butName)
        {
            Playback.Wait(3000);

            var group = windowInstence.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinButton>(new {Name = butName});
            var editControlcollection = Actions.GetControlCollection(editControl);

            return editControlcollection;
        }

        public static void ClickOnContinueBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ReportToAndBillingInfoWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Continue >");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnBackBtn()
        {
            Playback.Wait(3000);
            var windowInstence = ReportToAndBillingInfoWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "< Back");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }
    }
}