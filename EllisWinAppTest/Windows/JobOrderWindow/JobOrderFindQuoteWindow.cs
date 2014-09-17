using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;


namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class JobOrderFindQuoteWindow : AppContext
    {
        private static UITestControl GetCreateJobOrderWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Create New JobOrder" });
            return joborderWindow;
        }
        

        public static void EnterJobOrderFindQuoteData(DataRow data)
        {
            var getJobOrderWindow = GetCreateJobOrderWindowProperties();

            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            {

                Factory.SetMaskedText(getJobOrderWindow, "txtJobSiteZip", data.ItemArray[9].ToString());

            }

            //ClickOnButton("GO");
            MouseActions.ClickButton(getJobOrderWindow,"btnGo");
            //Enter data in dropdown fields
            Playback.Wait(2000);
            if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbState",
                data.ItemArray[10].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
                DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbCity",
                data.ItemArray[11].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                DropDownActions.SelectDropdownByText(getJobOrderWindow, "_cboPostalCodes",
                data.ItemArray[9].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                DropDownActions.SelectDropdownByText(getJobOrderWindow, "cmbCounty",
                data.ItemArray[12].ToString());

            MouseActions.ClickButton(getJobOrderWindow, "btnGo");
            

        }

        public static void ClickOnButton(string btnName)
        {
            Factory.ClickOnButton(GetCreateJobOrderWindowProperties(), btnName);
            //MouseActions.ClickButton(GetCreateJobOrderWindowProperties(), btnName);
        }


    }
}
