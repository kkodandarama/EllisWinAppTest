using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Data;

namespace EllisWinAppTest.Windows.JobOrderWindow
{
    internal class BasicJobInformationWindow : AppContext
    {
        private static UITestControl GetBasicJobInformationWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Create New JobOrder"});
            return joborderWindow;
        }

        private static UITestControlCollection GetBasicJobInformationWindowEditControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetBasicJobInformationWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinEdit>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }

        private static UITestControlCollection GetBasicJobInformationWindowDropdownControlCollection()
        {
            Playback.Wait(3000);
            var joborderWindow = GetBasicJobInformationWindowProperties();
            var group = joborderWindow.Container.SearchFor<WinGroup>(new {Name = ""});
            var editControl = group.Container.SearchFor<WinComboBox>(new {Name = ""});
            var editControlcollection = Actions.GetControlCollection(editControl);
            return editControlcollection;
        }


        public static void EnterBasicJobInformationWindowData(DataRow data)
        {
            // Get Dropdown controls and enter data
            var joborderWindow = GetBasicJobInformationWindowProperties();
            //var getBasicJobInformationControlCollection = GetBasicJobInformationWindowDropdownControlCollection();

            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                DropDownActions.SelectDropdownByText(joborderWindow,
                    BasicJobInformaionConstants.OrderPlacedBy, data.ItemArray[13].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
                DropDownActions.SelectDropdownByText(joborderWindow, BasicJobInformaionConstants.JobDuty,
                     data.ItemArray[14].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[22].ToString()))
                DropDownActions.SelectDropdownByText(joborderWindow,
                     BasicJobInformaionConstants.JobSitesList, data.ItemArray[22].ToString());
            if (!string.IsNullOrEmpty(data.ItemArray[24].ToString()))
                DropDownActions.SelectDropdownByText(joborderWindow,
                     BasicJobInformaionConstants.JobSiteContactList, data.ItemArray[24].ToString());


            // Get Edit controls and enter data
            var getBasicJobInformationWindow = GetBasicJobInformationWindowProperties();

            var poNumber = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.PoNumber);
            poNumber.SetFocus();
            SendKeys.SendWait("");
            SendKeys.SendWait(data.ItemArray[15].ToString());

            var custProjName = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.CustomerProjectName);
            Actions.SetText(custProjName, data.ItemArray[19].ToString());

            var custCostCentre = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.CustomerCostCentre);
            Actions.SetText(custCostCentre, data.ItemArray[18].ToString());

            var jobSiteName = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.JobSiteName);
            Actions.SetText(jobSiteName, data.ItemArray[23].ToString());

            var addLine1 = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.AddressLine1);
            Actions.SetText(addLine1, data.ItemArray[25].ToString());

            var addLine2 = Actions.GetWindowChild(getBasicJobInformationWindow, BasicJobInformaionConstants.AddressLine2);
            Actions.SetText(addLine2, data.ItemArray[26].ToString());



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
            var windowInstence = GetBasicJobInformationWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Continue >");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnBackBtn()
        {
            Playback.Wait(3000);
            var windowInstence = GetBasicJobInformationWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "< Back");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }

        public static void ClickOnCancelJobOrderBtn()
        {
            Playback.Wait(3000);
            var windowInstence = GetBasicJobInformationWindowProperties();
            var butColloction = GetButtonColloction(windowInstence, "Cancel Job Order");

            foreach (var control in butColloction)
            {
                MouseActions.Click(control);
            }
        }


        public class BasicJobInformaionConstants
        {
            //Combo Box/List Box
            public const string OrderPlacedBy = "cmbOrderPlacedBy";
            public const string JobDuty = "cmbJobDuty";
            public const string JobSitesList = "cmbJobSiteName";
            public const string JobSiteContactList = "cmbJobSiteContact";

            //Edit Fields
            public const string JoEffectiveDate = "dtpEffectiveDate";
            public const string JoExpirationDate = "dtpExpitarionDate";
            public const string PoNumber = "txtPONumber";
            public const string CustomerProjectName = "txtCustomerProjectName";
            public const string CustomerCostCentre = "txtCustomerCostCentre";
            public const string JobSiteName = "txtJobSiteName";
            public const string AddressLine1 = "txtAddressLine1";
            public const string AddressLine2 = "txtAddressLine2";
            //public const string CustProjectName = "cmbOrderPlacedBy";
            //public const string JobSiteContactName = "";

            //Radio Button
            public const string WorkTicketPref = "rdo";

            //Check Box
            public const string CodPayment = "chk";

            //Buttons
        }
    }
}
