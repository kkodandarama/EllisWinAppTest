using System;
using System.Diagnostics;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Collections.Generic;
using System.Globalization;

namespace EllisWinAppTest.Windows.CustomerWindow
{
    public class CustomerProfileWindow : AppContext
    {
        #region Window Properties

        public static UITestControl GetCustomerProfileWindowProperties()
        {
            var ccustomerWindow =
                App.Container.SearchFor<WinWindow>(new {Name = Globals.CustomerName + " - Customer Profile"});
            GlobalWindows.CustomerProfileWindow = ccustomerWindow;
            return ccustomerWindow;
        }

        public static UITestControl GetCustomerContactInfoWindowProperties()
        {
            var cContactInfoWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Customer Contact Information"});
            GlobalWindows.CustomerContactInfoWindow = cContactInfoWindow;
            return cContactInfoWindow;
        }

        public static UITestControl GetWorkersCompWindowProperties()
        {
            var cWCompWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Customer Workers\' Compensation Codes"});
            GlobalWindows.CustomerWorkerComp = cWCompWindow;
            return cWCompWindow;
        }

        private static UITestControl GetSelectOrgWindowProperties()
        {
            var cOrgWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Select Organization"});
            GlobalWindows.CustomerWorkerComp = cOrgWindow;
            return cOrgWindow;
        }

        private static UITestControl GetAddNoteWindowProperties()
        {
            var addNoteWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Add Note"});
            GlobalWindows.CustomerAddNote = addNoteWindow;
            return addNoteWindow;
        }

        private static UITestControl TaxExcemptionWindowProperties()
        {
            var taxWindow =
                App.Container.SearchFor<WinWindow>(new {Name = "Tax Exemptions"});
            return taxWindow;
        }

        private static UITestControl GetUnSavedChangedWindowProperties()
        {
            var window =
               App.Container.SearchFor<WinWindow>(new { Name = "Un-Saved Changes" });
            return window;
        }

        private static UITestControl PrintCOIWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Print Certificate of Insurance"});
            return window;
        }

        private static UITestControl COIPreviewWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Certificate of Insurance Report"});
            GlobalWindows.COIPreviewWindow = window;
            return window;
        }

        private static UITestControl GetDOFWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Customer Document On File"});
            GlobalWindows.DocumentOnFileWindow = window;
            return window;
        }

        private static UITestControl GetDOFFindFileWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Customer Document on File - Find File"});
            GlobalWindows.DocumentOnFileFindFileWindow = window;
            return window;
        }

        private static UITestControl GetQuoteProfileWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Quote Profile"});
            GlobalWindows.QuoteProfileWindow = window;
            return window;
        }

        private static UITestControl GetCreateQuoteWizardProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Create Quote Wizard"});
            GlobalWindows.QuoteProfileWindow = window;
            return window;
        }

        private static UITestControl GetUnsavedChangesWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Un-Saved Changes" });
            return window;
        }

        private static UITestControl GetNoQuotingAllowedWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Create Quote" });
            return window;
        }

        private static UITestControl GetChangeCustomerOnCopiedWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Change Customer On Copied View" });
            return window;
        }

        private static UITestControl GetJobDutiesWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Select Job Duties"});
            return window;
        }

        private static UITestControl GetExceptionCodeWindow()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Exception Comp Code Warning"});
            return window;
        }

        private static UITestControl GetZoneMatrixWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Zone Comp-Code Matrixs"});
            return window;
        }

        private static UITestControl GetAddNewJobSiteWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Add New Job Site"});
            return window;
        }

        private static UITestControl GetWarningWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Warning"});
            return window;
        }

        private static UITestControl GetAddressWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new {Name = "Address"});
            return window;
        }

        private static UITestControl GetAddModifierWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Add/Edit Modifier" });
            return window;
        }

        private static UITestControl GetTooManyResultsWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Too Many Results" });
            return window;
        }

        public static WinWindow GetCustomerInvoiceWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Customer Invoice" });
            return window;
        }

        private static UITestControl GetJobOrderProfileWindowProperties(string number)
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = number+"-Job Order Profile" });
            GlobalWindows.JobOrderProfileWindow = window;
            return window;
        }

        public static UITestControl GetEditCompCodeModifierWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Edit Comp Code Modifier" });
            GlobalWindows.JobOrderProfileWindow = window;
            return window;
        }

        private static UITestControl GetApplyCCPaymentWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Apply Credit Card Payment" });
            return window;
        }

        private static UITestControl GetPrintDialogWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Print" });
            GlobalWindows.DialogWindow = window;
            return window;
        }

        private static UITestControl GetEllisExceptionInfoWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Ellis Exception Information" });
            return window;
        }

        private static UITestControl GetExistingCustomerFoundWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Existing Customer Found" });
            return window;
        }

        private static UITestControl GetChangeCustomerStatusWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Change Customer Status" });
            return window;
        }

        private static UITestControl GetCustomerLinkingWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Customer Linking" });
            return window;
        }

        private static UITestControl GetCustomerDeLinkingWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Delink Customer Parent" });
            return window;
        }

        private static UITestControl GetCustomerCreditLimitWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Change Customer Credit Limit" });
            return window;
        }

        private static UITestControl GetRequestCreditWindowProperties()
        {
            var window =
                App.Container.SearchFor<WinWindow>(new { Name = "Request Credit" });
            return window;
        }

        #endregion

        public static bool VerifyCustomerProfileWindowDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            return cProfileWindow.Exists;
        }

        private static bool VerifyExistingCustomerWindowDisplayed()
        {
            var window = GetExistingCustomerFoundWindowProperties();
            return window.Exists;
        }

        public static void SelectTab(string tabName)
        {
            var tabs = GetTabs();
            var tabList = tabs.Container.SearchFor<WinWindow>(new {ControlName = CProfileControls.CProfileDetailsTab});
            var selectedTab = tabList.Container.SearchFor<WinTabPage>(new {Name = tabName});
            MouseActions.Click(selectedTab);
        }

        public static void SelectCustomerSegmentation(string type)
        {
            var window = GetCustomerProfileWindowProperties();
            DropDownActions.SelectDropdownByText(window, CProfileControls.CustomerSegment, type);
        }

        public static bool VerifyCustomerSegment(string type)
        {
            var text = DropDownActions.GetCurrentItem(GetCustomerProfileWindowProperties(),
                CProfileControls.CustomerSegment);

            return (text.Equals(type));
        }

        public static bool VerifyProfileDefaults()
        {
            //var tabs = GetTabs();
            //var nameDBA = tabs.Container.SearchFor<WinWindow>(new {ControlName = CProfileControls.NameDBAWin});
            var winInst = GetCustomerProfileWindowProperties();
            var custName = Actions.GetWindowChild(winInst, "CaseTitleLabel");

            return custName.Exists;
        }
       
        //public static void ClickContactsTab()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();
        //    var cprofileTabs = Actions.GetWindowChild(cProfileWindow, CProfileControls.CProfileDetailsTab);
        //    var contact = cprofileTabs.Container.SearchFor<WinTabPage>(new { Name = CCTabConstants.Contacts });
        //    MouseActions.Click(contact);
        //}

        public static void SelectFromProfileDetailsTab(string tabName)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var cprofileTabs = Actions.GetWindowChild(cProfileWindow, CProfileControls.CProfileDetailsTab);
            var pfTab = cprofileTabs.Container.SearchFor<WinTabPage>(new { Name = tabName });
            MouseActions.Click(pfTab);
        }

        public static void ClickAddContactButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.AddContact);

            //var child = Actions.GetWindowChild(cProfileWindow, CProfileControls.AddContact);
            //var addContactBtn =
            //    child.Container.SearchFor<WinButton>(new {Name = CProfileControls.AddContactButton});
            //Mouse.Click(addContactBtn);
        }

        public static bool VerifyActiveCheckDisabled()
        {
            var cContactWindow = GetCustomerContactInfoWindowProperties();
            var active = Actions.GetWindowChild(cContactWindow, CProfileControls.ActiveCheck);
            return (active.Enabled);
        }

        public static bool VerifyNoteTypeOptionsDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var noteType = Actions.GetWindowChild(cProfileWindow, CProfileControls.NoteType);

            return noteType.Enabled;
        }

        public static void EnterContactDetails(DataRow data)
        {
            var cContactWindow = GetCustomerContactInfoWindowProperties();
            Globals.CustomerContact = Generator.GenerateNewName(data.ItemArray[3].ToString());

            Actions.SetText(cContactWindow, CProfileControls.ContactName, Globals.CustomerContact);
            Actions.SetText(cContactWindow, CProfileControls.Title, data.ItemArray[4].ToString());
            Actions.SetText(cContactWindow, CProfileControls.Site, data.ItemArray[5].ToString());

            var address = Actions.GetWindowChild(cContactWindow, CProfileControls.Address);
            DropDownActions.SelectDropdownByText(address, "Test");

            Actions.SetText(cContactWindow, CProfileControls.PhoneNumber, data.ItemArray[6].ToString());

            Playback.Wait(2000);
            Actions.SetCheckBox(cContactWindow, CProfileControls.AccountManager, data.ItemArray[7].ToString());

            SelectContactRoles(data.ItemArray[8].ToString());

            Actions.SetCheckBox(cContactWindow, CProfileControls.AutomatedEmails, data.ItemArray[7].ToString());

            SelectEmailTypes(data.ItemArray[10].ToString());

            var emailOptions = Actions.GetWindowChild(cContactWindow, CProfileControls.EmailOptions);
            ListActions.SelectFromList((WinList) emailOptions, data.ItemArray[11].ToString());
        }

        private static void SelectEmailTypes(string data)
        {
            string control;
            var cContactWindow = GetCustomerContactInfoWindowProperties();
            switch (data)
            {
                case "All":
                    control = "All orders";
                    break;
                case "Orders Placed":
                    control = "Orders placed by this contact";
                    break;
                default:
                    control = "All orders";
                    break;
            }

            Actions.SelectRadioButton(cContactWindow, control);
        }

        private static void SelectContactRoles(string data)
        {
            var contactInfo = GetCustomerContactInfoWindowProperties();
            //var tableName = Actions.GetWindowChild(contactInfo, CProfileControls.Types);
            //var table = (WinTable) tableName;
            //var row = table.Container.SearchFor<WinRow>(new {Name = "KeyValuePair`2 row 4"});
            //var cell = row.Container.SearchFor<WinCell>(new {Name = "Selected"});
            var cell = TableActions.SelectCellFromTable(contactInfo, CProfileControls.Types, "KeyValuePair`2 row 4",
                "Selected");
            MouseActions.Click(cell);
        }

        private static UITestControl GetTabs()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var children = cProfileWindow.GetChildren();
            return children[3];
        }

        public static void ClickSave()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.Save);
        }

        public static void ClickCancel()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            try
            {
                TitlebarActions.ClickClose((WinWindow)cProfileWindow);
            }
            catch
            {
                Actions.SendAltF4();
            }
        }

        public static bool VerifyNewContactAdded(string name)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.SearchResultGrid);
            var table = (WinTable) tableName;
            for (var i = 1; i < 4; i++)
            {
                var row = table.Container.SearchFor<WinRow>(new {Name = CProfileControls.SearchResultRow + i});
                var cell = row.Container.SearchFor<WinCell>(new {Name = "Name"});
                if (cell.Value.Equals(name))
                    return true;
            }
            return false;
        }

        public static void SelectStatusFilter(string select)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var statusWindow = Actions.GetWindowChild(cProfileWindow, CProfileControls.StatusFilter);
            var status = (WinComboBox) statusWindow;
            DropDownActions.SelectDropdownByText(status, select);
        }

        public static bool SelectFirstCustomerFromTable()
        {
            try
            {
                var cell = TableActions.SelectCellFromTable(EllisWindow, CProfileControls.CustomersGrid,
                "CustomerQuoteSummaryDomain row 1", "Customer Name");
                //cell.SetFocus();
                Globals.CustomerName = cell.Value;
                //Mouse.DoubleClick(cell);
                MouseActions.DoubleClick(cell);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            
        }

        public static bool SelectRandomCustomerFromTable()
        {
            try
            {
            var rand = Generator.GenerateRandomNumber(2);
            var cell = TableActions.SelectCellFromTable(EllisWindow, CProfileControls.CustomersGrid,
                    "CustomerQuoteSummaryDomain row " + rand, "Customer Name"); ;
            Globals.CustomerName = cell.Value;
            MouseActions.DoubleClick(cell);
                return true;
        }
            catch (Exception)
            {
                return false;
            }
        }

        public static void SelectFirstCustomerContactFromTable()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cell = TableActions.SelectCellFromTable(cProfileWindow, CProfileControls.SearchResultGrid,
                CProfileControls.SearchResultRow + 1, "Name");
            Globals.CustomerContact = cell.Value;
            MouseActions.DoubleClick(cell);
        }

        public static void ClickWorkersCompButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.WorkersComp);
        }

        public static void AddWorkersCompToProfile()
        {
            if (!VerifyAddCodeButtonEnabled())
                ClickAssociatedCodeCancel();
            //else
            //{
            //    ClickAssociatedAddCode();
            //}
        }

        public static bool VerifySelectOrganizationButtonEnabled()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var selectBtn = Actions.GetWindowChild(cProfileWindow, CProfileControls.AddOrganization);
            return selectBtn.Enabled;
        }

        public static void ClickSelectOrganizationButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.AddOrganization);
        }

        public static void SelectOrganizationDetails()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();

            var coordinator = Actions.GetWindowChild(selectOrgWindow, CProfileControls.User);
            DropDownActions.SelectDropdownByText(coordinator, "Test EllisCSR");
        }

        public static void SelectAllOrganizationDetails()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();

            var orgType = Actions.GetWindowChild(selectOrgWindow, CProfileControls.OrganizationType);
            Factory.SelectDropdownByText(orgType, "Branch");

            var org = Actions.GetWindowChild(selectOrgWindow, CProfileControls.Organization);
            Factory.SelectDropdownByText(org, "1111");

            var coordinator = Actions.GetWindowChild(selectOrgWindow, CProfileControls.User);
            Factory.SelectDropdownByText(coordinator, "Test EllisCSR");
        }

        public static void ClickSelectButton()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();
            MouseActions.ClickButton(selectOrgWindow, CProfileControls.Select);
        }

        public static bool VerifyOrganizationDetailsUpdated()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var coordinator = Actions.GetWindowChild(cProfileWindow, CProfileControls.AccountCoordinator);
            var coordin = (WinText) coordinator;

            return coordin.DisplayText.Equals("Test EllisCSR");
        }

        public static bool VerifyQuotingRulesDisabled(string ruleType)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var type = Actions.GetWindowChild(cProfileWindow, ruleType);
            return !type.Enabled;
        }

        private static bool VerifyAddCodeButtonEnabled()
        {
            var cWorkComp = GetWorkersCompWindowProperties();
            var addCodeWin = Actions.GetWindowChild(cWorkComp, CProfileControls.WorkersCompAddCode);
            return addCodeWin.Enabled;
        }

        public static void ClickAssociatedCodeCancel()
        {
            var cWorkComp = GetWorkersCompWindowProperties();
            MouseActions.ClickButton(cWorkComp, CProfileControls.WorkersCompCancel);
        }

        public static void ClickAssociatedAddCode()
        {
            var cWorkComp = GetWorkersCompWindowProperties();
            MouseActions.ClickButton(cWorkComp, CProfileControls.WorkersCompAddCode);
        }

        public static bool VerifyModifierSetToOne()
        {
            var prof = GetEditCompCodeModifierWindowProperties();
            var cell = TableActions.SelectCellFromTable(prof, CProfileControls.ModifiersGrid,
                "WorkersCompensationCustomerModifierProfileDomain row 1", "Modifier");
            var value = cell.Value.Substring(0, cell.Value.IndexOf('.', 0));
            return value.Equals("1");
        }

         public static bool VerifyAssociatedCodeTypeSetToPrimary()
        {
             var prof = GetEditCompCodeModifierWindowProperties();
             var drop = Actions.GetWindowChild(prof, CProfileControls.CodeAssociationType);
             var dropdown = (WinComboBox) drop;
             var text = dropdown.SelectedItem;

             return text.Equals("Primary");
        }

        public static void ClickAddModifierWindow()
        {
            var prof = GetEditCompCodeModifierWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.AddModifier);
        }

        public static bool VerifyAddModifierWindowDisplayed()
        {
            var prof = GetAddModifierWindowProperties();
            return prof.Exists;
        }

        public static void CloseAddModifierWindow()
        {
            var prof = GetAddModifierWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.Cancel);
        }

        public static void ClickAddNewNoteButton(string type)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var noteType = Actions.GetWindowChild(cProfileWindow, CProfileControls.NoteType);
            var noteTypeCombo = (WinComboBox) noteType;
            DropDownActions.SelectDropdownByText(noteTypeCombo, type);

            ClickAddNoteButton();
        }

        public static void ClickAddNoteButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.AddNote);
        }

        public static void AddNewNote(string noteText)
        {
            var addNoteWindow = GetAddNoteWindowProperties();
            Actions.SetText(addNoteWindow, CProfileControls.Notes, noteText);
        }

        public static void ClickOnSaveAndCloseButton()
        {
            var addNoteWindow = GetAddNoteWindowProperties();
            MouseActions.ClickButton(addNoteWindow, "_btnSave");
        }

        public static bool VerifyNewNoteDisplayedInGrid(string note)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var cell = TableActions.SelectCellFromTable(cProfileWindow, CProfileControls.NotesGrid,
                CProfileControls.NotesTable, "Note");
            var noteDisplayed = cell.Value;

            return noteDisplayed.Contains(note);
        }

        public static void EnterDate(string datetype, string date)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            UITestControl dateWindow = null;

            switch (datetype)
            {
                case "From":
                    dateWindow = Actions.GetWindowChild(cProfileWindow, CProfileControls.FromDate);
                    break;
                case "To":
                    dateWindow = Actions.GetWindowChild(cProfileWindow, CProfileControls.ToDate);
                    break;
            }

            var comboBox = (WinComboBox) dateWindow;
            MouseActions.Click(comboBox);
            for (var i = 0; i < 12; i++)
            {
                SendKeys.SendWait("{BACKSPACE}");
                Playback.Wait(200);
            }

            for (var i = 0; i < 5; i++)
            {
                SendKeys.SendWait("{DELETE}");
                Playback.Wait(200);
            }
            //SendKeys.SendWait("{HOME}");
            Playback.Wait(1500);
            SendKeys.SendWait(date);
        }

        public static void ClickGoButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.Go);
        }

        public static void ClickClearButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.Clear);
        }

        public static bool VerifyEmptyGridDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            try
            {
                var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.NotesGrid);
                var table = (WinTable)tableName;
                return table.Cells.Count == 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #region Credit Info methods

        public static void SelectCreditInfoSubTab(string tabName)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var cLimitTabs = Actions.GetWindowChild(cProfileWindow, CProfileControls.CCreditInfoTab);
            var clTab = cLimitTabs.Container.SearchFor<WinTabPage>(new {Name = tabName});
            MouseActions.Click(clTab);
        }

        public static bool VerifyCreditStatus(string data)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cStatus = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditStatus);
            var status = (WinText) cStatus;
            Globals.Temp = status.Name;
            return status.Name.Equals(data);
        }

        public static bool VerifyCreditTerms()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cTerms = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditTerms);
            var terms = (WinText) cTerms;
            Globals.Temp = terms.Name;
            return terms.Name == "Credit" || terms.Name == "COD";
        }

        public static bool VerifyCurrentCreditLimitDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cLimit = Actions.GetWindowChild(cProfileWindow, CProfileControls.CurrentCreditLimit);
            var limit = (WinText) cLimit;
            Globals.Temp = limit.Name;
            return limit.Enabled;
        }

        public static bool VerifyCurrentCreditLimitDisplayed(string amount)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cLimit = Actions.GetWindowChild(cProfileWindow, CProfileControls.CurrentCreditLimit);
            var limit = (WinText)cLimit;

            return limit.Name.Equals(amount);
        }

        public static bool VerifyCreditLimitWarnAndLockoutDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cWarn = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditLimitWarn);
            var warn = (WinEdit)cWarn;

            var cLockout = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditLimitLockout);
            var lockout = (WinEdit)cLockout;

            return warn.Enabled && lockout.Enabled;
        }

        public static bool VerifyCreditLimitNotesDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditLimitHistoryTable);
            var table = (WinTable) tableName;
            return table.Exists;
        }

        public static bool ClickUpdateCreditLimit()
        {
            try
            {
                var window = GetCustomerProfileWindowProperties();
                MouseActions.ClickButton(window, CProfileControls.UpdateCreditLimit);
                return Actions.GetWindowChild(window, CProfileControls.UpdateCreditLimit).Enabled;
            }
            catch (Exception)
            {
                return false;
            }
            
        }

        public static void AddNewCreditLimit(string amount)
        {
            var window = GetCustomerCreditLimitWindowProperties();

            Actions.SetText(window, CProfileControls.NewCredit, amount);
            Actions.SetText(window, CProfileControls.ReasonForChange, "ELLIS NAPS TEST");
            MouseActions.ClickButton(window, CProfileControls.ApplyChange);
        }
        
        public static bool VerifyCreditStatusDropDown(string data)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cStatus = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditStatusDDown);
            var status = (WinComboBox) cStatus;
            Globals.Temp = status.SelectedItem;
            return status.SelectedItem.Equals(data);
        }

        public static bool VerifyCreditTermsDropDown()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cTerms = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditTermsDDown);
            var terms = (WinComboBox) cTerms;
            Globals.Temp = terms.SelectedItem;
            return terms.SelectedItem == "Credit" || terms.Name == "COD";
        }

        //public static bool VerifyCreditOnApplicationFileIsChecked()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();
        //    var cof = Actions.GetWindowChild(cProfileWindow, CProfileControls.CreditApplicationOnFile);
        //    var cOnFile = (WinCheckBox) cof;
        //    return cOnFile.Checked;
        //}

        public static bool VerifyDateSignedOnDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var csigned = Actions.GetWindowChild(cProfileWindow, CProfileControls.ApplicationDate);
            var signedOn = (WinComboBox) csigned;
            return signedOn.Exists;
        }

        public static bool VerifyBusinessLicensesDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.BusineesLicenses);
            var table = (WinTable) tableName;
            return table.Exists;
        }

        #endregion

        #region Billing Methods

        //public static void ClickSelectInvoicingManagmentButton()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();

        //    var btnName = Actions.GetWindowChild(cProfileWindow, CProfileControls.SelectOrganization);
        //    var btn = btnName.Container.SearchFor<WinButton>(new { Name = "Select" });
        //    Mouse.Click(btn);
        //}

        public static void UncheckJobOrderOwner()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();
            Actions.SetCheckBox(selectOrgWindow, CProfileControls.JobOrderOwner, "False");
        }

        public static void CheckSingleOrganization()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();
            Actions.SetCheckBox(selectOrgWindow, CProfileControls.OrganizationOwner, "True");
        }

        public static bool VerifyBillingCoordinator()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var billingCoordinator = Actions.GetWindowChild(cProfileWindow, CProfileControls.BillingCoordinator);
            var bCoordinator = (WinText) billingCoordinator;
            return bCoordinator.DisplayText.Contains("Test EllisCSR");
        }

        //public static bool VerifyEBilling()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();
        //    var ebilling = Actions.GetWindowChild(cProfileWindow, CProfileControls.EBilling);
        //    var cOnFile = (WinCheckBox)ebilling;
        //    if (!cOnFile.Checked) return false;
        //    Console.WriteLine("Customer has opted for e-billing");
        //    return true;
        //}

        public static void CheckConsolidatedInvoicing()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            Actions.SetCheckBox(cProfileWindow, CProfileControls.ConsolidatedInvoicing, "True");
        }

        //public static bool VerifyConsolidatedInvoicingChecked()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();
        //    var cInv = Actions.GetWindowChild(cProfileWindow, CProfileControls.ConsolidatedInvoicing);
        //    var cOnFile = (WinCheckBox)cInv;
        //    return cOnFile.Checked;
        //}

        public static void CheckOvertimeBilling()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            Actions.SetCheckBox(cProfileWindow, CProfileControls.OvertimeBilling, "True");
        }

        //public static bool VerifyOvertimeBillingSelected()
        //{
        //    var cProfileWindow = GetCustomerProfileWindowProperties();
        //    var oBill = Actions.GetWindowChild(cProfileWindow, CProfileControls.OvertimeBilling);
        //    var cOnFile = (WinCheckBox)oBill;
        //    return cOnFile.Checked;
        //}

        public static void ClickViewTaxExceptions()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.TaxExceptions);

            //var tax = Actions.GetWindowChild(cProfileWindow, CProfileControls.TaxExceptions);
            //var taxBtn = tax.Container.SearchFor<WinButton>(new {Name = "View Tax Exemptions"});
            //Mouse.Click(taxBtn);
        }

        public static bool VerifyExceptionStatusSetToAll()
        {
            var excemptionWindow = TaxExcemptionWindowProperties();
            var excempt = Actions.GetWindowChild(excemptionWindow, CProfileControls.ExcemptionStatus);
            var eStatus = (WinComboBox) excempt;
            Globals.Temp = eStatus.SelectedItem;
            return eStatus.SelectedItem.Equals("All");
        }

        public static void ClickExcemptionWindowCancel()
        {
            var excemptionWindow = TaxExcemptionWindowProperties();
            var btn = excemptionWindow.Container.SearchFor<WinButton>(new {Name = "Cancel"});
            btn.SetFocus();
            MouseActions.ClickButton(excemptionWindow, "_btnCancel");
        }

        #endregion

        #region Collections Methods

        public static void ClickInvoicingOrgSelectBtn()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.InvOrgSelect);
            //var invOrgBtn = Actions.GetWindowChild(cProfileWindow, CProfileControls.InvOrgSelect);
            //var btn = invOrgBtn.Container.SearchFor<WinButton>(new {Name = "Select"});
            //Mouse.Click(btn);
        }

        public static void UncheckInvoicingOrganization()
        {
            var ellisWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "Ellis"));
            var selectOrgWindow = ellisWindow.FindFirst(TreeScope.Children, new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "Select Organization"));
            var invoicingCheckbox = selectOrgWindow.FindFirst(TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(AutomationElement.AutomationIdProperty, "_ugbxJobOrderOwner"));

            if (invoicingCheckbox != null)
                Mouse.Click(Factory.GetAutomationElementMiddlePoint(invoicingCheckbox));
        }

        public static bool VerifyAgingThesholdsSectionDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var branchOverride = Actions.GetWindowChild(cProfileWindow, CProfileControls.OverrideBranch);
            var percentThreshold = Actions.GetWindowChild(cProfileWindow, CProfileControls.WarnPercentThreshold);
            var thresholdDays = Actions.GetWindowChild(cProfileWindow, CProfileControls.WarnThresholdDays);
            var percentLockout = Actions.GetWindowChild(cProfileWindow, CProfileControls.PercentLockout);
            var thresholdLockout = Actions.GetWindowChild(cProfileWindow, CProfileControls.ThresholdDaysLockout);

            return branchOverride.Exists && percentThreshold.Exists 
                && thresholdDays.Exists && percentLockout.Exists && thresholdLockout.Exists;
        }

        public static bool VerifyPrimaryOrgUnitDisplayed()
        {
            var selectOrgWindow = GetSelectOrgWindowProperties();
            var selectOrg = Actions.GetWindowChild(selectOrgWindow, CProfileControls.PrimaryOrg);
            var org = (WinText) selectOrg;
            return org.DisplayText.Contains("1111");
        }

        public static void ClickPrintCOIButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            //var coiBtn = Actions.GetWindowChild(cProfileWindow, CProfileControls.PrintCOI);
            //var btn = coiBtn.Container.SearchFor<WinButton>(new {Name = "Print COI"});
            //Mouse.Click(btn);
            MouseActions.ClickButton(cProfileWindow, CProfileControls.PrintCOI);
        }

        public static void ClickPrintButton()
        {
            var coiWindow = PrintCOIWindowProperties();
            //var printBtn = Actions.GetWindowChild(coiWindow, CProfileControls.PrintCOI);
            //var btn = printBtn.Container.SearchFor<WinButton>(new {Name = "Print"});
            //MouseActions.Click(btn);
            MouseActions.ClickButton(coiWindow, CProfileControls.PrintCOI);
        }

        public static bool VerifyPrintPreviewWindowDisplayed()
        {
            var coiPreview = COIPreviewWindowProperties();
            return coiPreview.Enabled;
        }

        public static void ClosePPWindow()
        {
            TitlebarActions.ClickClose(GlobalWindows.COIPreviewWindow);
        }

        public static bool SelectFinanceCharges(string data)
        {
            string control;
            var cProfileWindow = GetCustomerProfileWindowProperties();
            switch (data)
            {
                case "Yes":
                    control = "Yes";
                    break;
                case "No":
                    control = "No";
                    break;
                default:
                    control = "Yes";
                    break;
            }

            try
            {
                var result = Actions.SelectRadioButton(cProfileWindow, control);
                return result;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Servicing Methods

        public static void SelectServicingInnerTab(string controlName)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var tabWorkSpace = Actions.GetWindowChild(cProfileWindow, CProfileControls.TabWorkspace);
            var tabPage = tabWorkSpace.Container.SearchFor<WinTabPage>(new {Name = controlName});
            MouseActions.Click(tabPage);
        }

        public static bool VerifyCheckboxDisabled(string name)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            string control = null;
            switch (name)
            {
                case "Account Manager":
                    control = CProfileControls.AccountManagerPermitted;
                    break;
                case "Assigned To Branch":
                    control = CProfileControls.AssignedToBranch;
                    break;
                case "Authorized to Order":
                    control = CProfileControls.AuthorizedToOrder;
                    break;
                case "Consolidated Invoicing":
                    control = CProfileControls.ConsolidatedInvoicing;
                    break;
                case "Overtime Billing":
                    control = CProfileControls.OvertimeBilling;
                    break;
                case "EBilling":
                    control = CProfileControls.EBilling;
                    break;
                case "Credit Application":
                    control = CProfileControls.CreditApplicationOnFile;
                    break;
                case "Active Check":
                    control = CProfileControls.ActiveCheck;
                    break;
                case "Alert User":
                    control = CProfileControls.AlertUser;
                    break;
                case "ECommerce Access":
                    control = CProfileControls.ECommercePermit;
                    break;
                case "Branch Permitted":
                    control = CProfileControls.BranchPermitted;
                    break;
                case "Purchase Order Required":
                    control = CProfileControls.PurchaseOrder;
                    break;
                case "Drug Test":
                    control = CProfileControls.DrugTest;
                    break;
                case "Background Check":
                    control = CProfileControls.BackgroundCheck;
                    break;
                case "Bahavioral Survey":
                    control = CProfileControls.BehavorialSurvey;
                    break;
            }

            var chkBox = Actions.GetWindowChild(cProfileWindow, control);
            return !chkBox.Enabled;
        }

        public static bool SelectServiceCoordinator(string name)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var servCord = Actions.GetWindowChild(cProfileWindow, CProfileControls.ServiceCoordinator);
            var status = (WinComboBox) servCord;
            DropDownActions.SelectDropdownByText(status, name);

            return status.SelectedItem.Equals(name);
        }

        public static bool VerifyServiceCoordinatorEnabled()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var servCord = Actions.GetWindowChild(cProfileWindow, CProfileControls.ServiceCoordinator);
            return servCord.Exists;
        }


        public static void ClickCancelButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.CancelButton);
        }

        public static bool SetPOFormat(string data)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var nameDBA = Actions.GetWindowChild(cProfileWindow, CProfileControls.POFormat);
            var nameEdit = (WinEdit) nameDBA;
            Actions.SetText(nameEdit, data);

            var value = nameEdit.GetProperty("Value");

            return value.Equals(Globals.CustomerName);
        }

        public static void ClickAddFileButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.AddFile);
        }

        public static void FindCustomerDocumentOnFile()
        {
            var findDOF = GetDOFWindowProperties();
            var type = Actions.GetWindowChild(findDOF, CProfileControls.DocumentType);
            var docType = (WinComboBox) type;
            DropDownActions.SelectDropdownByText(docType, "Contract");

            type = Actions.GetWindowChild(findDOF, CProfileControls.OperationalCompany);
            docType = (WinComboBox) type;
            DropDownActions.SelectDropdownByText(docType, "Labor Ready Central Inc");

            MouseActions.ClickButton(findDOF, CProfileControls.FindFile);
        }

        public static void ClickCloseOnCustomerDOF()
        {
            var openFileDialog = AutomationElement.RootElement.FindFirst(
                TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(
                    AutomationElement.NameProperty, "Customer Document on File - Find File"));
            
            var openFileDialogCancelButton = openFileDialog.FindFirst(
                TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(
                    AutomationElement.NameProperty, "Cancel"));

            for (int i = 0; i < 3; i++)
            {
                Actions.SendTab();
            }
            Actions.SendEnter();
        }

        public static void ClickCloseOnDOF()
        {
            var findDOF = GetDOFWindowProperties();
            MouseActions.ClickButton(findDOF, CProfileControls.CancelButton);
        }

        #endregion

        #region Quotes Methods

        public static void EnterQuoteSearchCriteria()
        {
            var cProfile = GetCustomerProfileWindowProperties();

            var status = Actions.GetWindowChild(cProfile, CProfileControls.QuoteStatus);
            var cmbBox = (WinComboBox) status;
            DropDownActions.SelectDropdownByText(cmbBox, "Accepted");

            var state = Actions.GetWindowChild(cProfile, CProfileControls.State);
            cmbBox = (WinComboBox) state;
            DropDownActions.SelectDropdownByText(cmbBox, "All");

            var city = Actions.GetWindowChild(cProfile, CProfileControls.City);
            cmbBox = (WinComboBox) city;
            DropDownActions.SelectDropdownByText(cmbBox, "All");

            var zip = Actions.GetWindowChild(cProfile, CProfileControls.ZipCode);
            cmbBox = (WinComboBox) zip;
            DropDownActions.SelectDropdownByText(cmbBox, "All");

            try
            {
                var child = Actions.GetWindowChild(cProfile, CProfileControls.ChildList);
                cmbBox = (WinComboBox)child;
                DropDownActions.SelectDropdownByText(cmbBox, "All Linked Accounts");
            }
            catch (Exception)
            {
                //Suppress any exception here
            }

            MouseActions.ClickButton(cProfile, CProfileControls.Gobtn);
        }

        public static bool SelectFirstQuoteFromResults()
        {
            Playback.Wait(3000);
            try
            {
                var cProfile = GetCustomerProfileWindowProperties();
                var cell = TableActions.SelectCellFromTable(cProfile, CProfileControls.QuoteSummary,
                    "QuoteGridDisplayDomain row 2", "Quoted By");
                Globals.QuotedBy = cell.Value;
                MouseActions.DoubleClick(cell);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            
        }

        public static void ClickCopyQuote()
        {
            var cProfile = GetCustomerProfileWindowProperties();
           MouseActions.ClickButton(cProfile, CProfileControls.CopyQuote);
        }

        public static bool ClickChangeQuoteCustomerName()
        {
            try
            {
                var prof = GetCreateQuoteWizardProperties();
                MouseActions.ClickButton(prof, CProfileControls.ChangeQuoteCustomer);
                return true;
            }
            catch (Exception)
            {
                return false;
            } 
            
        }

        public static void SelectFromQuoteProfileTab(string tabName)
        {
            var qProfileWindow = GetQuoteProfileWindowProperties();
            var qprofileTabs = Actions.GetWindowChild(qProfileWindow, CProfileControls.TabQuoteProfile);
            var qfTab = qprofileTabs.Container.SearchFor<WinTabPage>(new {Name = tabName});
            MouseActions.Click(qfTab);
        }

        public static void SelectNewCustomerNameAfterSearch(string name)
        {
            var prof = GetChangeCustomerOnCopiedWindowProperties();
            Actions.SetText(prof, CProfileControls.CustomerName, name);

            MouseActions.ClickButton(prof, CProfileControls.GoButton);

            try
            {
                var prof2 = GetTooManyResultsWindowProperties();

                MouseActions.ClickButton(prof2, CProfileControls.OkBtn2);
            }
            catch (Exception)
            {
                //Suppress any Exception here
            }

            MouseActions.ClickButton(prof, CProfileControls.OkBtn);
        }

        public static bool VerifyNoQuotingAllowedWindowDisplayed()
        {
            var prof = GetNoQuotingAllowedWindowProperties();
            return prof.Exists;
        }

        public static void CloseNoQuotingAllowedWindow()
        {
            try
            {
                var prof = GetNoQuotingAllowedWindowProperties();
                MouseActions.ClickButton(prof, CProfileControls.OkBtn2);
            }
            catch (Exception)
            {
                //Suppress any expection here
            }
        }

        public static void CloseUnsavedChangesWindow()
        {
            try
            {
                var prof = GetUnsavedChangesWindowProperties();
                MouseActions.ClickButton(prof, CProfileControls.OkBtn2);
            }
            catch (Exception)
            {
                //Suppress any expection here
            }
        }

        public static bool VerifyCreateNewQuoteWindowDisplayed()
        {
            var qProfileWindow = GetCreateQuoteWizardProperties();
            return qProfileWindow.Exists;
        }

        public static bool VerifyQuoteStatusDisplayed()
        {
            var qProfileWindow = GetQuoteProfileWindowProperties();
            var qStatusLabel = Actions.GetWindowChild(qProfileWindow, CProfileControls.QuoteStatusLabel);
            var qstatus = (WinText) qStatusLabel;
            return qstatus.Name != null;
        }

        public static bool VerifyPricingMatrixDisplayed()
        {
            var qProfileWindow = GetQuoteProfileWindowProperties();
            var tableName = Actions.GetWindowChild(qProfileWindow, CProfileControls.PricingMatrix);
            var table = (WinTable) tableName;

            return table.Exists;
        }

        public static bool VerifyQuoteHistoryGridTable()
        {
            var qProfileWindow = GetQuoteProfileWindowProperties();
            var tableName = Actions.GetWindowChild(qProfileWindow, CProfileControls.QuoteHistoryGrid);
            var table = (WinTable) tableName;

            return table.Exists;
        }

        #endregion

        #region Create New Quote Methods

        public static void ClickCreateNewQuoteButton()
        {
            var cProfile = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfile, CProfileControls.CreateNewQuoteBtn);
        }

        public static void CreateNewQuote(DataRow data)
        {
            var quoteWindow = GetCreateQuoteWizardProperties();
            var jobDuty = GetJobDutiesWindowProperties();
            var exceptionWindow = GetExceptionCodeWindow();
            var zoneMatrix = GetZoneMatrixWindowProperties();

            Actions.SetText(quoteWindow, CProfileControls.RequestorName, data.ItemArray[3].ToString());
            Actions.SetText(quoteWindow, CProfileControls.RequestorPhone, data.ItemArray[4].ToString());

            var date = Actions.GetWindowChild(quoteWindow, CProfileControls.EffectiveDate);
            var cmbBox = (WinComboBox) date;
            var dateData = Actions.TrimDate(data.ItemArray[5].ToString());
            DropDownActions.SelectDropdownByText(cmbBox, DateTime.Today.ToString("d"));

            Actions.SetText(quoteWindow, CProfileControls.PONumber, data.ItemArray[6].ToString());
            MouseActions.ClickButton(quoteWindow, CProfileControls.ItemInclude);

            Playback.Wait(500);
            Actions.SendText("20.00");
            Playback.Wait(1500);

            MouseActions.ClickButton(quoteWindow, CProfileControls.Continue);
            quoteWindow.WaitForControlReady(6000);

            //Actions.SendAltF4();
            if(GetUnsavedChangesWindowProperties().Exists)
                MouseActions.ClickButton(GetUnSavedChangedWindowProperties(), "_CancelButton");

            MouseActions.ClickButton(quoteWindow, CProfileControls.JobDuties);
            MouseActions.ClickButton(jobDuty, CProfileControls.AddJobDuty);
            MouseActions.ClickButton(jobDuty, CProfileControls.OkBtn);
            Playback.Wait(2000);

            var states = Actions.GetWindowChild(quoteWindow, CProfileControls.States);
            DropDownActions.SelectDropdownByText(states, "WASHINGTON");

            MouseActions.ClickButton(quoteWindow, CProfileControls.AddZone);
            MouseActions.ClickButton(quoteWindow, CProfileControls.ExceptionCodes);

            try
            {
                MouseActions.ClickButton(exceptionWindow, CProfileControls.OkBtn2);
            }
            catch (Exception)
            {
                //supress any exceptions here
            }

            var des = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description);
            DropDownActions.SelectDropdownByText(des, "WA");

            var des2 = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description2);
            DropDownActions.SelectDropdownByText(des2, "Manufacturing");

            var des3 = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description3);
            DropDownActions.SelectDropdownByText(des3, "Electronic or Scientific Assembly and Printing");

            var comp = Actions.GetWindowChild(zoneMatrix, CProfileControls.CompCode);
            DropDownActions.SelectDropdownByText(comp, String.Empty);

            //var tableName = Actions.GetWindowChild(zoneMatrix, CProfileControls.CompCode);
            //var table = (WinTable)tableName;

            //var row = table.Container.SearchFor<WinRow>(new { Name = "QuoteWorkersCompensation row 1" });
            //var row = TableActions.SelectRowFromTable(zoneMatrix, CProfileControls.CompCode,
            //    "QuoteWorkersCompensation row 1");
            //Mouse.Click(row);

            Playback.Wait(200);
            SendKeys.SendWait("{DOWN}");
            
            MouseActions.ClickButton(exceptionWindow, CProfileControls.AddCompCode);
            MouseActions.ClickButton(exceptionWindow, CProfileControls.OkBtn);

            //var price = Actions.GetWindowChild(quoteWindow, CProfileControls.PricingMethod);
            Factory.SetMaskedText(quoteWindow, CProfileControls.PricingMethod, data.ItemArray[12].ToString());
            //var editableTextbox = price.GetChildren().ToArray()[0] as WinEdit;
            //editableTextbox.Text = "_________.00";
            
            //Actions.SetText(price, data.ItemArray[12].ToString());

            var pMethod = Actions.GetWindowChild(quoteWindow, CProfileControls.PricingMethodCmb);
            DropDownActions.SelectDropdownByText(pMethod, data.ItemArray[11].ToString());

            var bPay = Actions.GetWindowChild(quoteWindow, CProfileControls.BasisForPay);
            DropDownActions.SelectDropdownByText(bPay, data.ItemArray[13].ToString());

            MouseActions.ClickButton(quoteWindow, CProfileControls.AddDefaultsButton);

            var cell = TableActions.SelectCellFromTable(quoteWindow, CProfileControls.GridPricingMatrix,
                "QuoteGridDisplayDomain row 1", "Bill Rate");
            Globals.QuotedBy = cell.Value;

            MouseActions.ClickButton(quoteWindow, CProfileControls.Continue);
            Playback.Wait(2500);

            MouseActions.ClickButton(quoteWindow, CProfileControls.SaveAndClose);
            Playback.Wait(1500);
        }

        public static void CreateNewQuoteWithDefaultData()
        {
            var quoteWindow = GetCreateQuoteWizardProperties();
            var jobDuty = GetJobDutiesWindowProperties();
            var exceptionWindow = GetExceptionCodeWindow();
            var zoneMatrix = GetZoneMatrixWindowProperties();

            Actions.SetText(quoteWindow, CProfileControls.RequestorName,  "Test_Requestor");
            Actions.SetText(quoteWindow, CProfileControls.RequestorPhone, "5646789534");

            var date = Actions.GetWindowChild(quoteWindow, CProfileControls.EffectiveDate);
            var cmbBox = (WinComboBox)date;
            var dateData = Actions.TrimDate("4/29/2014");
            DropDownActions.SelectDropdownByText(cmbBox, dateData);

            Actions.SetText(quoteWindow, CProfileControls.PONumber, "DefaultTest");
            MouseActions.ClickButton(quoteWindow, CProfileControls.ItemInclude);

            Playback.Wait(500);
            Actions.SendText("20.00");
            Playback.Wait(1500);

            MouseActions.ClickButton(quoteWindow, CProfileControls.Continue);
            Playback.Wait(1500);

            Actions.SendAltF4();

            MouseActions.ClickButton(quoteWindow, CProfileControls.JobDuties);
            MouseActions.ClickButton(jobDuty, CProfileControls.AddJobDuty);
            MouseActions.ClickButton(jobDuty, CProfileControls.OkBtn);
            Playback.Wait(2000);

            var states = Actions.GetWindowChild(quoteWindow, CProfileControls.States);
            DropDownActions.SelectDropdownByText(states, "WASHINGTON");

            MouseActions.ClickButton(quoteWindow, CProfileControls.AddZone);
            MouseActions.ClickButton(quoteWindow, CProfileControls.ExceptionCodes);

            try
            {
                MouseActions.ClickButton(exceptionWindow, CProfileControls.OkBtn2);
            }
            catch (Exception)
            {
                //supress any exceptions here
            }

            var des = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description);
            DropDownActions.SelectDropdownByText(des, "WA");

            var des2 = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description2);
            DropDownActions.SelectDropdownByText(des2, "Manufacturing");

            var des3 = Actions.GetWindowChild(zoneMatrix, CProfileControls.Description3);
            DropDownActions.SelectDropdownByText(des3, "Electronic or Scientific Assembly and Printing");

            var comp = Actions.GetWindowChild(zoneMatrix, CProfileControls.CompCode);
            DropDownActions.SelectDropdownByText(comp, "");

            var row = TableActions.SelectRowFromTable(zoneMatrix, CProfileControls.CompCode,
                "QuoteWorkersCompensation row 1");
            Mouse.Click(row);

            Playback.Wait(200);
            SendKeys.SendWait("{DOWN}");

            MouseActions.ClickButton(exceptionWindow, CProfileControls.AddCompCode);
            MouseActions.ClickButton(exceptionWindow, CProfileControls.OkBtn);

            var price = Actions.GetWindowChild(quoteWindow, CProfileControls.PricingMethod);
            for (int i = 0; i < 10; i++)
            {
                Playback.Wait(200);
                SendKeys.SendWait("{BACKSPACE}");
                SendKeys.SendWait("{DELETE}");
            }

            Playback.Wait(200);
            SendKeys.SendWait("{HOME}");
            Actions.SetText(price, "35");

            var pMethod = Actions.GetWindowChild(quoteWindow, CProfileControls.PricingMethodCmb);
            DropDownActions.SelectDropdownByText(pMethod, "Bill Rate");

            var bPay = Actions.GetWindowChild(quoteWindow, CProfileControls.BasisForPay);
            DropDownActions.SelectDropdownByText(bPay, "Zone Pay");

            MouseActions.ClickButton(quoteWindow, CProfileControls.AddDefaultsButton);

            var cell = TableActions.SelectCellFromTable(quoteWindow, CProfileControls.GridPricingMatrix,
                "QuoteGridDisplayDomain row 1", "Bill Rate");
            Globals.QuotedBy = cell.Value;
            MouseActions.DoubleClick(cell);

            MouseActions.ClickButton(quoteWindow, CProfileControls.Continue);
            Playback.Wait(1500);

            MouseActions.ClickButton(quoteWindow, CProfileControls.SaveAndClose);
            Playback.Wait(1500);
        }

        #endregion

        #region Create New Quote Methods

        public static void CloseWarningWindow()
        {
            try
            {
                var warningWindow = GetWarningWindowProperties();
                MouseActions.ClickButton(warningWindow, CProfileControls.OkBtn);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Current exception found is : "+ex);
            }
        }

        public static void CloseAddressWindow()
        {
            try
            {
                var warningWindow = GetAddressWindowProperties();
                MouseActions.ClickButton(warningWindow, "_btnOk");
            }
            catch (Exception)
            {
                //suppress any exception here
            }
        }

        public static void CloseWarningWindowWindow()
        {
            try
            {
                var warningWindow = GetWarningWindowProperties();
                MouseActions.ClickButton(warningWindow, "&OK");
            }
            catch (Exception)
            {
                //suppress any exception here
            }
        }

        #endregion

        #region Create New Job Order Methods

        public static void ClickCreateNewJobOrderButton()
        {
            var cProfile = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfile, CProfileControls.CreateNewJobOrder);
        }

        public static void EnterJobOrderSearchCriteria()
        {
            var cProfile = GetCustomerProfileWindowProperties();

            var cList = Actions.GetWindowChild(cProfile, CProfileControls.ChildList);
            DropDownActions.SelectDropdownByText(cList, "All Linked Accounts");

            var joStatus = Actions.GetWindowChild(cProfile, CProfileControls.JobOrderStatus);
            DropDownActions.SelectDropdownByText(joStatus, "All");

            var sate = Actions.GetWindowChild(cProfile, CProfileControls.JoStates);
            DropDownActions.SelectDropdownByText(sate, "WA");

            var city = Actions.GetWindowChild(cProfile, CProfileControls.Cities);
            DropDownActions.SelectDropdownByText(city, "TACOMA");

            var zip = Actions.GetWindowChild(cProfile, CProfileControls.PostalCode);
            DropDownActions.SelectDropdownByText(zip, "98404");

            MouseActions.ClickButton(cProfile, CProfileControls.GoBtn);
        }

        public static void SelectFirstJobOrderFromResults()
        {
            var cProfile = GetCustomerProfileWindowProperties();
            var cell = TableActions.SelectCellFromTable(cProfile, CProfileControls.JobOrderSummary,
                "CustomerJobOrderSummaryDomain row 1", "Job Order #");
            Globals.Temp = cell.Value;
            MouseActions.DoubleClick(cell);
        }

        public static bool VerifyBasicJobInfoDisplayed()
        {
            var joProfile = OpenJobOrder.JobOrderProfileWindowProperties();
            return joProfile.Exists;
        }

        public static void SelectJOSubTab(string tabName)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cprofileTabs = Actions.GetWindowChild(cProfileWindow, CProfileControls.JobOrderTab);
            var pfTab = cprofileTabs.Container.SearchFor<WinTabPage>(new {Name = tabName});
            MouseActions.Click(pfTab);
        }

        public static void ClickAddNewJobSiteButton()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(cProfileWindow, CProfileControls.AddNewJobSiteBtn);
        }

        public static bool VerifyAddNewJobSiteWindowDispllayed()
        {
            var prof = GetAddNewJobSiteWindowProperties();
            return prof.Exists;
        }

        public static void CreateNewJobSite(DataRow data)
        {
            var newJS = GetAddNewJobSiteWindowProperties();
            ClickAddJobSiteTab(CCTabConstants.JobSiteDetails);

            Actions.SetCheckBox(newJS, CProfileControls.ActiveCheck, data.ItemArray[3].ToString());

            Globals.Temp = Generator.GenerateNewName(data.ItemArray[4].ToString());
            Actions.SetText(newJS, CProfileControls.JobSiteName, Globals.Temp);

            Actions.SetText(newJS, CProfileControls.JSAddress1, data.ItemArray[5].ToString());
            Actions.SetText(newJS, CProfileControls.JSAddress2, data.ItemArray[6].ToString());

            var country = Actions.GetWindowChild(newJS, CProfileControls.JSCountry);
            DropDownActions.SelectDropdownByText(country, data.ItemArray[7].ToString());

            var state = Actions.GetWindowChild(newJS, CProfileControls.JSState);
            DropDownActions.SelectDropdownByText(state, data.ItemArray[8].ToString());

            var zip = Actions.GetWindowChild(newJS, CProfileControls.JSZip);
            DropDownActions.SelectDropdownByText(zip, data.ItemArray[9].ToString());

            var city = Actions.GetWindowChild(newJS, CProfileControls.JSCity);
            DropDownActions.SelectDropdownByText(city, data.ItemArray[10].ToString());

            var county = Actions.GetWindowChild(newJS, CProfileControls.JSCounty);
            DropDownActions.SelectDropdownByText(county, data.ItemArray[11].ToString());

            ClickAddJobSiteTab(CCTabConstants.ReportToContact);
            Playback.Wait(1000);

            Actions.SetText(newJS, CProfileControls.ReportToContactName, data.ItemArray[12].ToString());
            Actions.SetText(newJS, CProfileControls.ReportToContactTitle, data.ItemArray[13].ToString());
            Actions.SetText(newJS, CProfileControls.ReportDepartment, data.ItemArray[14].ToString());
            Actions.SetText(newJS, CProfileControls.ReportEmail, data.ItemArray[15].ToString());

            MouseActions.ClickButton(newJS, CProfileControls.JSSave);

            CloseAddressWindow();
            Playback.Wait(2000);
        }

        public static bool VerifyNewJobSiteCreated(DataRow data, string siteName)
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var state = Actions.GetWindowChild(cProfileWindow, CProfileControls.JSStates);
            DropDownActions.SelectDropdownByText(state, data.ItemArray[8].ToString());

            var city = Actions.GetWindowChild(cProfileWindow, CProfileControls.JSCities);
            DropDownActions.SelectDropdownByText(city, data.ItemArray[10].ToString());

            Playback.Wait(1000);

            var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.JSSummary);
            var table = (WinTable) tableName;
            return table.Exists;

            //for (int i = 1; i < 15; i++)
            //{
            //    var row = table.Container.SearchFor<WinRow>(new { Name = "CustomerJobSiteDomain row " +i});
            //    var cell = row.Container.SearchFor<WinCell>(new { Name = "Job Site Name" });

            //    if (cell.Value.Equals(siteName))
            //        return true;
            //}
            //return false;
        }

        public static bool VerifyWorkersDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var tableName = Actions.GetWindowChild(cProfileWindow, CProfileControls.WorkerGrid);
            var table = (WinTable) tableName;
            return table.Exists;
        }

        private static void ClickAddJobSiteTab(string tabName)
        {
            var newJS = GetAddNewJobSiteWindowProperties();
            var ultraGrid = Actions.GetWindowChild(newJS, CProfileControls.UltraTab);
            var uTab = ultraGrid.Container.SearchFor<WinTabPage>(new {Name = tabName});
            MouseActions.Click(uTab);
        }

        #endregion

        #region Invoices and Payments Methods

        public static void SelectInvSubTab(string tabName)
        {
            Thread.Sleep(1000);
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var children = cProfileWindow.GetChildren();
            var tabs= children[3];
            var tabList = tabs.Container.SearchFor<WinWindow>(new { ControlName = CProfileControls.InvTab });
            var selectedTab = tabList.Container.SearchFor<WinTabPage>(new { Name = tabName });
            MouseActions.Click(selectedTab);
        }



        public static void SearchAndSelectFirstInvoiceFromGrid()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var status = Actions.GetWindowChild(cProfileWindow, "_cbStatus");
            DropDownActions.SelectDropdownByText(status, "All");

            MouseActions.ClickButton(cProfileWindow, CProfileControls.Go);

            SelectFirstInvoiceFromGrid();
        }

        public static void SelectFirstInvoiceFromGrid()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();
            var cell = TableActions.SelectCellFromTable(cProfileWindow, CProfileControls.InvGrid,
               "CustomerInvoiceSummaryDomain row 1", "Invoice Number");
            Globals.Temp = cell.Value;
            Mouse.DoubleClick(cell);
        }

        public static void ClickRequestCredit()
        {
            var window = GetCustomerInvoiceWindowProperties();
            var temp = Actions.GetWindowChild(window, CProfileControls.RequestAmount);
            Globals.Amount = temp.Name;
            MouseActions.ClickButton(window, CProfileControls.RequestCredit);
        }

        public static void EnterCreditRequestData()
        {
            var window = GetRequestCreditWindowProperties();
            DropDownActions.SelectDropdownByText(window, CProfileControls.CreditType, "Customer Sales Credit");
            Actions.SetText(window, CProfileControls.CreditComments, "ELLIS NAPS TEST");
            Actions.SetText(window, CProfileControls.CreditAmount, "10.00");

            MouseActions.ClickButton(window, CProfileControls.RequestCreditButton);
        }

        public static bool VerifyPendingCreditRequestAmount(string amount)
        {
            var window = GetCustomerInvoiceWindowProperties();
            var temp = Actions.GetWindowChild(window, CProfileControls.RequestAmount);
            string temp2 = Globals.Amount.Substring(1);
            var amt = Convert.ToDecimal(temp2) + Convert.ToDecimal(amount);
            string something = "$" +Convert.ToString(amt);
            return temp.Name.Equals(something);
        }

        public static void SearchForUnPaidInvoices()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            var status = Actions.GetWindowChild(cProfileWindow, "_cbStatus");
            DropDownActions.SelectDropdownByText(status, "Un-Paid");

            MouseActions.ClickButton(cProfileWindow, CProfileControls.Go);
        }

        public static bool VerifyUnpaidInvoicesDisplayed()
        {
            var cProfileWindow = GetCustomerProfileWindowProperties();

            try
            {
                var cell = TableActions.SelectCellFromTable(cProfileWindow, CProfileControls.InvGrid,
               "CustomerInvoiceSummaryDomain row 1", "Invoice Number");
                Globals.Temp = cell.Value;
            }
            catch (Exception)
            {
                //Suppress all exceptions here
            }

            return Globals.Temp != null;
        }

        public static bool VerifyCustomerInvoiceNumberDisplayed(string number)
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var lbl = Actions.GetWindowChild(prof, CProfileControls.InvoiceNumber);
            return lbl.Name.Equals("Invoice # "+number);
        }

        public static void SelectCustomerInvoiceSubTab(string name)
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var cprofileTabs = Actions.GetWindowChild(prof, CProfileControls.UltraTab);
            var pfTab = cprofileTabs.Container.SearchFor<WinTabPage>(new { Name = name });
            MouseActions.Click(pfTab);
        }

        public static void SearchForTransactionHistory()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var dd = Actions.GetWindowChild(prof, CProfileControls.PaymentType);
            DropDownActions.SelectDropdownByText(dd, "All");

            dd = Actions.GetWindowChild(prof, CProfileControls.GetHistoryFor);
            DropDownActions.SelectDropdownByText(dd, "All Invoices");

            dd = Actions.GetWindowChild(prof, CProfileControls.SinceDate);
            DropDownActions.SelectDropdownByText(dd, "03/05/2012");

            MouseActions.ClickButton(prof, CProfileControls.HistoryFilter);
        }

        public static bool VerifyBalanceDueDisplayed()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var bal = Actions.GetWindowChild(prof, CProfileControls.BalanceAmount);
            return bal.Exists;
        }

        public static void ClickOnJobOrderNumber()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var link = Actions.GetWindowChild(prof, CProfileControls.JobOrderDetail);
            Globals.Temp = link.FriendlyName;
            MouseActions.Click(link);
        }

        public static void ClickOnJobOrderNumberLink()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var link = Actions.GetWindowChild(prof, CProfileControls.JobOrderNumber);
            Globals.Temp = link.FriendlyName;
            MouseActions.Click(link);
        }

        public static bool VerifyJobOrderWindowDisplayed()
        {
            var prof = GetJobOrderProfileWindowProperties(Globals.Temp);
            return prof.Exists;
        }

        public static bool ApplyCcPaymentEnabled()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            var btn = Actions.GetWindowChild(prof, CProfileControls.ApplyCC);
            return btn.Exists;
        }

        public static void ClickApplyCCPayment()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.ApplyCC);
        }

        public static bool VerifyApplyCCPaymentWindowDisplayed()
        {
            var prof = GetApplyCCPaymentWindowProperties();
            return prof.Exists;
        }

        public static void EnterCreditCardDetails(DataRow data)
        {
            var window = GetApplyCCPaymentWindowProperties();

            Actions.SetText(window, CProfileControls.CardHolderName, data.ItemArray[3].ToString());
            Actions.SetText(window, CProfileControls.CardAddress1, data.ItemArray[4].ToString());
            Actions.SetText(window, CProfileControls.CardCity, data.ItemArray[5].ToString());
            DropDownActions.SelectDropdownByText(window, CProfileControls.CardAddress1, data.ItemArray[6].ToString());
            Actions.SetText(window, CProfileControls.CardZip, data.ItemArray[7].ToString());
            DropDownActions.SelectDropdownByText(window, CProfileControls.CardType, data.ItemArray[8].ToString());
            Actions.SetText(window, CProfileControls.CardNumber, data.ItemArray[9].ToString());
            Actions.SetText(window, CProfileControls.CardDate, data.ItemArray[10].ToString());
            Actions.SetText(window, CProfileControls.CardCVV, data.ItemArray[11].ToString());
            Actions.SetText(window, CProfileControls.CardAmount, data.ItemArray[13].ToString());
        }

        public static void ClickProcessPayment()
        {
            var window = GetApplyCCPaymentWindowProperties();
            MouseActions.ClickButton(window, CProfileControls.ProcessPayment);
        }

        public static void CloseApplyCCPaymentWindow()
        {
            var prof = GetApplyCCPaymentWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.Cancel);
        }

        public static void ClickOnReprintOriginalButton()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.ReprintOriginal);
        }

        public static void ClickOnPrintAdjustedButton()
        {
            var prof = GetCustomerInvoiceWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.PrintAdjusted);
        }

        public static bool VerifyPrintDialogWindowDisplayed()
        {
            var prof = GetPrintDialogWindowProperties();
            return prof.Exists;
        }

        public static bool VerifyEllisExceptionWindowDisplayed()
        {
            var prof = GetEllisExceptionInfoWindowProperties();
            if (prof.Exists)
            {
                MouseActions.ClickButton(prof, CProfileControls.ExceptionOk);
                return true;
            }
            return false;
        }

        public static bool VerifyQuoteProfileWindowDisplayed()
        {
            var window = GetQuoteProfileWindowProperties();
            return window.Exists;
        }

        public static void SearchForAllInvoices()
        {
            ChangeSinceDate("02/02/2013");
            ClickHistoryFilter();
        }

        public static void ClickHistoryFilter()
        {
            var prof = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(prof, CProfileControls.HistoryFilter);
        }

        public static bool VerifySearchResultsDisplayed()
        {
            var prof = GetCustomerProfileWindowProperties();
            
            string value = null;

            try
            {
                var cell = TableActions.SelectCellFromTable(prof, CProfileControls.PaymentHistoryGrid,
                "PaymentHistorySummaryDomain row 1", "Invoice Number");
                value = cell.Value;
            }
            catch (Exception)
            {
                //suppress any exception here
            }
            return !string.IsNullOrEmpty(value);
        }

        private static void ChangeSinceDate(string date)
        {
            var prof = GetCustomerProfileWindowProperties();
            //var dateContainer = prof.Container.SearchFor<WinComboBox>(new {Name = CProfileControls.SinceDate});
            var child = Actions.GetWindowChild(prof, CProfileControls.SinceDate);
            child.SetFocus();
            SendKeys.SendWait("^A");
            SendKeys.SendWait("{DELETE}");
            SendKeys.SendWait(date);
        }

        #endregion

        #region Management Methods

        public static void ClickChangeStatus()
        {
            var window = GetCustomerProfileWindowProperties();
            MouseActions.ClickButton(window, CProfileControls.ChangeStatus);
        }

        public static void ChangeStatus(string status)
        {
            ClickChangeStatus();
            var window = GetChangeCustomerStatusWindowProperties();
            DropDownActions.SelectDropdownByText(window, CProfileControls.DDChangeStatus, status);
            DropDownActions.SelectDropdownByText(window, CProfileControls.ReasonForStatusChange, "Account Management");
            Keyboard.SendKeys("{TAB}");
            Keyboard.SendKeys("ELLIS NAPS TEST");
            //Actions.SetText(window, CProfileControls.ReasonForChangeTxt, "ELLIS NAPS TEST");
            MouseActions.ClickButton(window, CProfileControls.Save);
        }

        public static bool VerifyStatusChanged(string status)
        {
            var window = GetChangeCustomerStatusWindowProperties();
            var cStatus = Actions.GetWindowChild(window, CProfileControls.AccountStatus);
            var label = (WinText)cStatus;
            return label.Name.Equals(status);
        }

        public static void ClickCustomerLinking()
        {
            var window = GetCustomerLinkingWindowProperties();
            MouseActions.ClickButton(window, CProfileControls.CustomerLinking);
        }

        public static void EnterParentCustomerNumber()
        {
            var window = GetCustomerLinkingWindowProperties();
            Actions.SetText(window, CProfileControls.ParentNumber, "000674722");
        }

        public static void ClickValidate()
        {
            var window = GetCustomerLinkingWindowProperties();
            MouseActions.ClickButton(window, CProfileControls.ValidateParent);
        }

        public static void DelinkParentCustomer()
        {
            var window = GetCustomerLinkingWindowProperties();
            var window2 = GetCustomerDeLinkingWindowProperties();
            Actions.SetText(window, CProfileControls.ParentNumber, "");
            MouseActions.ClickButton(window, CProfileControls.ValidateParent);
            MouseActions.ClickButton(window, CProfileControls.SaveCustomerLinking);

            MouseActions.ClickButton(window2, CProfileControls.OkBtn2);
        }

        public static bool VerifyInheritanceOptionsEnabled()
        {
            var window = GetCustomerLinkingWindowProperties();
            Actions.SetCheckBox(window, CProfileControls.CreditCheck, "True");
            Actions.SetCheckBox(window, CProfileControls.ParentQuote, "True");
            Actions.SetCheckBox(window, CProfileControls.ParentCollections, "True");
            Actions.SetCheckBox(window, CProfileControls.OrderingRules, "True");
            Actions.SetCheckBox(window, CProfileControls.InvoiceAndBilling, "True");

            if(VerifyCustomerLnkingCheckBox("Credit") == "True")
                if (VerifyCustomerLnkingCheckBox("Quote") == "True")
                    if (VerifyCustomerLnkingCheckBox("Collections") == "True")
                        if (VerifyCustomerLnkingCheckBox("Ordering") == "True")
                            if (VerifyCustomerLnkingCheckBox("Billing") == "True")
                                return true;
            return false;
        }

        public static void ClickSaveCustomerLinking()
        {
            var window = GetCustomerLinkingWindowProperties();
            MouseActions.ClickButton(window, CProfileControls.SaveCustomerLinking);
        }

        private static string VerifyCustomerLnkingCheckBox(string name)
        {
            var cProfileWindow = GetCustomerLinkingWindowProperties();
            string control = null;
            switch (name)
            {
                case "Credit":
                    control = CProfileControls.CreditCheck;
                    break;
                case "Quote":
                    control = CProfileControls.ParentQuote;
                    break;
                case "Collections":
                    control = CProfileControls.ParentCollections;
                    break;
                case "Ordering":
                    control = CProfileControls.OrderingRules;
                    break;
                case "Billing":
                    control = CProfileControls.InvoiceAndBilling;
                    break;
            }
            var chkBox = Actions.GetWindowChild(cProfileWindow, control);
            var set= chkBox.GetProperty("Checked").ToString();
            return set;
        }

        public static bool VerifyChainLinkDisplayed()
        {
            Thread.Sleep(3000);
            var window = GetCustomerProfileWindowProperties();
            try
            {
                var child = Actions.GetWindowChild(window, CProfileControls.ChainLink);
                if (child.Exists)
                    return true;
            }
            catch (Exception)
            {
                //Suppress all exceptions here
            }
            return false;
        }

        public static bool VerifyQuotingCheckboxesEnabled()
        {
            var window = GetCustomerProfileWindowProperties();
            var management = window.Container.SearchFor<WinWindow>(new { ControlName = CCTabConstants.ManageQuotes });
            var rate = window.Container.SearchFor<WinWindow>(new { ControlName = CCTabConstants.RateAgreementWindow });

            return management.Enabled && rate.Enabled;
        }

        public static bool VerifyInvOrgDisplayed()
        {
            var window = GetCustomerProfileWindowProperties();
            var control = window.Container.SearchFor<WinWindow>(new {Name = CProfileControls.InvoiceOrganization});
            var invoiceReviewOrganization = Actions.GetWindowChild(window, "_ulblInvoiceReviewOrganization");
            return invoiceReviewOrganization.FriendlyName.Equals("1111 - PUYALLUP BRANCH");
        }

        #endregion

        #region Close Windows Methods

        public static void CloseCustomerProfileWindow()
        {
            try
            {
                TitlebarActions.ClickClose(GlobalWindows.CustomerProfileWindow);
            }
            catch
            {
                //Suppressing all exceptions here
            }
        }

        public static void CloseCustomerSearchResultsWindow()
        {
            try
            {
                TitlebarActions.ClickClose(GlobalWindows.CustomerSearchResultsWindow);
            }
            catch
            {
                Actions.SendAltF4();
            }
        }

        public static void CloseCustomerInvoiceWindow()
        {
            TitlebarActions.ClickClose(App.Container.SearchFor<WinWindow>(new { Name = "Customer Invoice" }));
        }

        public static void CloseQuoteProfileWindow()
        {
            TitlebarActions.ClickClose(GlobalWindows.QuoteProfileWindow);
        }

        #endregion

        private class CProfileControls
        {
            #region Profile Details Window Controls

            public const string CProfileDetailsTab = "_tabCustomerProfileDetails";
            public const string NameDBAWin = "_txtNameDBA";
            public const string AddContact = "_btnAddContact";
            public const string AddContactButton = "Add Contact";
            public const string StatusFilter = "_cmbStatusFilter";
            public const string CustomersGrid = "_grdCustomers";
            public const string CustomerSegment = "_cmbCustomerSegmentation";

            #endregion

            #region Add Workers Comp Controls

            public const string WorkersComp = "_btnShowWorkersComp";
            public const string WorkersCompCancel = "_btnWorkersCompCancel";
            public const string WorkersCompAddCode = "_btnWorkersCompAddCode";
            public const string AddOrganization = "_btnSelectOrganization";
            public const string AccountCoordinator = "_lblAccountCoordinator";

            public const string ModifiersGrid ="_grdModifiers";
            public const string CodeAssociationType ="_cbCompCodeAssociationType";
            public const string AddModifier ="_btnAddModifier";

            #endregion

            #region Add Organization Controls

            public const string OrganizationType = "_cboOrganizationType";
            public const string Organization = "_cboOrganization";
            public const string User = "_cboUser";
            public const string Cancel = "_btnCancel";
            public const string Select = "_btnSelect";

            #endregion

            public const string SearchResultGrid = "_grdSearchResult";
            public const string SearchResultRow = "CustomerContactSummaryDomain row ";
            public const string CancelBtn = "btnCancel";

            #region Add Contact Info Controls

            public const string ActiveCheck = "_chkActive";
            public const string ContactName = "_txtContactName";
            public const string Title = "_txtTitle";
            public const string Site = "_txtBranchDept";
            public const string Address = "_cbAddress";
            public const string PhoneNumber = "_txtPhoneNumber";
            public const string AccountManager = "_chkAccountManager";
            public const string Types = "_grdContactTypes";
            public const string AutomatedEmails = "_chkReceiveAutomatedEmails";
            public const string AllOrders = "_rbAllOrders";
            public const string EmailOptions = "_lbEmailAdviceOptions";

            #endregion

            #region Add Notes Controls

            public const string AddNote = "_btnAddNote";
            public const string NoteType = "_cbNoteType";
            public const string Notes = "_txtNotes";
            public const string Save = "_btnSave";
            public const string NotesGrid = "_gridNotes";
            public const string NotesTable = "ultraGrid1";
            public const string FromDate = "_dtFromDate";
            public const string ToDate = "_dtAsOfDate";
            public const string Go = "_btnGo";
            public const string Clear = "_btnClear";

            #endregion

            #region Credit Info Controls

            public const string UpdateCreditLimit = "_btnUpdateCreditLimit";
            public const string NewCredit = "_txtNewCreditLimit";
            public const string ReasonForChange = "_txtReasonforChange";
            public const string ApplyChange = "_btnApplyChange";

            public const string CreditLimitWarn = "_uneCreditLimitWarn";
            public const string CreditLimitLockout = "_uneCreditLimitLockout";

            public const string RequestCredit = "_btnRequestCredit";

            #endregion

            #region Credit Limit Controls

            public const string CCreditInfoTab = "_tabCustomerProfileCredit";
            public const string CreditStatus = "_lblCreditStatus";
            public const string CreditTerms = "_lblCreditTermsDisplay";
            public const string CurrentCreditLimit = "_lblCreditLimit";

            public const string CreditLimitHistoryTable = "_grdCreditLimitHistory";

            public const string CreditStatusDDown = "_cmbCreditStatus";
            public const string CreditTermsDDown = "_cmbCreditTerms";
            public const string CreditApplicationOnFile = "_chkCreditApplicationOnFile";
            public const string ApplicationDate = "_dttApplicationDate";

            public const string BusineesLicenses = "_grdBusinessLicenses";

            public const string CreditType ="_cboxType";
            public const string CreditComments = "_txteditorComments";
            public const string CreditAmount = "_txtAmount";
            public const string RequestCreditButton = "_btnRequestCredit";
            public const string RequestAmount = "_lblCreditRequestAmt";

            #endregion

            #region Billing Controls

            //public const string SelectOrganization = "_btnSelectOrganization";
            public const string JobOrderOwner = "_uchkbxJobOrderOwner";
            public const string OrganizationOwner = "_uchkbxOrganization";

            public const string InvoiceOrganization = "_ulblInvoiceReviewOrganization";
            public const string BillingCoordinator = "_ulblBillingCoordinator";

            public const string RedirectToCorporate = "_uchkbxRedirectInvoicesToCorporate";
            public const string EBilling = "_uchkbxEBilling";
            public const string EmailAdd = "_utxtedtEBillingEmailAddress";

            public const string ConsolidatedInvoicing = "_uchkbxUseConsolidatedInvoicing";
            public const string OvertimeBilling = "_uchkbxOvertimeBillingPermitted";

            public const string TaxExceptions = "_btnViewTaxExemptions";
            public const string ExcemptionStatus = "_cbExemptionStatus";

            #endregion

            #region Collections Controls

            public const string InvOrgSelect = "_btnPrimary";
            public const string InvoicingOrg = "_uchkbxInvoicingOrganization";
            public const string PrimaryOrg = "_lblPrimaryOrg";
            public const string PrintCOI = "_btnPrint";
            public const string COIToolStrip = "toolStrip1";
            public const string FinanceCharges = "_optFinanceCharges";
            public const string OverrideBranch = "_uceOverrideBranchThreshold";
            public const string WarnPercentThreshold = "_umePercentThresholdWarn";
            public const string WarnThresholdDays = "_umeThresholdDaysWarn";
            public const string PercentLockout = "_umePercentLockout";
            public const string ThresholdDaysLockout = "_umeThresholdDaysLockout";

            #endregion

            #region Servicing Controls

            public const string TabWorkspace = "TabWorkspace";
            public const string AccountManagerPermitted = "_onlyAccountManagerPermittedCheckEditor";
            public const string AssignedToBranch = "_assignedToBranchMayRepeatOrdersCheckEditor";
            public const string ServiceCoordinator = "_serviceCoordinatorComboEditor";
            public const string AuthorizedToOrder = "_useAuthorizedToOrderContactsOnlyCheckEditor";
            public const string AlertUser = "_alertUserToServicingRulesCheckEditor";
            public const string ECommercePermit = "_permittedToAccessECommerceSiteCheckEditor";
            public const string BranchPermitted = "_branchPermittedToContactForNewBusinessCheckEditor";
            public const string CancelButton = "_cancelButton";
            public const string SaveButton = "_saveButton";

            public const string PurchaseOrder = "_purchaseOrderRequiredCheckEditor";
            public const string POFormat = "_purchaseOrderFormatEditor";

            public const string DrugTest = "_drugTestCheckEditor";
            public const string BackgroundCheck = "_backgroundCheckRequiredCheckEditor";
            public const string BehavorialSurvey = "_behavioralSurveyRequiredCheckEditor";

            public const string AddFile = "_addFileButton";
            public const string DocumentType = "_documentTypesComboEditor";
            public const string OperationalCompany = "_operationalCompanyComboEditor";
            public const string FindFile = "_findFileButton";

            #endregion

            #region Credit Card Payment Controls

            public const string CardHolderName = "_txtCardHolderName";
            public const string CardAddress1 = "_txtAddress1";
            public const string CardAddress2 = "_txtAddress2";
            public const string CardCity = "_txtCity";
            public const string CardStates = "_cboStates";
            public const string CardZip = "_txtZipCode";
            public const string CardPhone = "_txtPhone";
            public const string CardFax = "_txtFaxNumber";
            public const string CardType = "_cboCardType";
            public const string CardCVV = "_txtCVV2";
            public const string CardNumber = "_txtCardNumber";
            public const string CardDate = "_txtExpDate";
            public const string CardAmount = "_txtAmount";

            public const string ProcessPayment = "_btnProcess";

            #endregion

            #region Quote Controls

            public const string QuoteStatus = "_cbQuoteStatus";
            public const string State = "_cmbState";
            public const string City = "_cmbCity";
            public const string ZipCode = "_cmbZipCode";
            public const string ChildList = "_cmbChildList";
            public const string Gobtn = "_btnGo";

            public const string QuoteSummary = "_grdQuoteSummary";
            public const string TabQuoteProfile = "_tabQuoteProfile";
            public const string QuoteStatusLabel = "lblQuoteStatus";

            public const string PricingMatrix = "_gridPricingMatrix";
            public const string QuoteHistoryGrid = "quoteHistoryGrid";

            public const string CopyQuote = "_btnCopyQuote";
            public const string ChangeQuoteCustomer = "btnChangeQuoteCustomer";
            public const string CustomerName ="txtCustomerName";
            public const string GoButton = "btnGo";

            #endregion

            #region Create New Quote Controls

            public const string CreateNewQuoteBtn = "_btnNewQuote";
            public const string RequestorName = "txtQuoteRequestorName";
            public const string RequestorPhone = "txtQuoteRequestorPhone";
            public const string EffectiveDate = "dtQuoteEffectiveDate";
            public const string PONumber = "txtCustomersPONumber";
            public const string ItemInclude = "btnInclude";
            public const string AdditionalItems = "UlGrdQuotesAdditionalChargableItems";

            public const string Continue = "ContinueButton";
            public const string JobDuties = "_btnSelectJobDuties";
            public const string AddJobDuty = "btnAddJobDuty";
            public const string OkBtn = "btnOk";

            public const string States = "_cboStates";
            public const string AddZone = "_btnAddZone";
            public const string ExceptionCodes = "_btnSelectExceptionCodes";
            public const string OkBtn2 = "_OKButton";

            public const string Description = "ulCbDescriptionAtLevel1";
            public const string Description2 = "ulCbDescriptionAtLevel2";
            public const string Description3 = "ulCbDescriptionAtLevel3";
            public const string CompCode = "ulCbCompCodes";
            public const string AddCompCode = "btnAdddCompCode";

            public const string PricingMethod = "_txtPricingMethodAmount";
            public const string AddDefaultsButton = "_btnUpdateByPricingAmount";
            public const string PricingMethodCmb = "_cboPricingMethod";
            public const string BasisForPay = "_cboBasisForPay";

            public const string GridPricingMatrix = "_gridPricingMatrix";

            public const string SaveAndClose = "FinishButton";

            #endregion

            #region Create New Job Order Controls

            public const string CreateNewJobOrder = "_btnNewJobOrder";
            public const string JobOrderStatus = "_cboJobOrderStatus";
            public const string JoStates = "_cmbStates";
            public const string Cities = "_cmbCities";
            public const string PostalCode = "_cmbPostalCode";
            public const string GoBtn = "btnFilter";

            public const string JobOrderSummary = "_grdJobOrderSummary";
            public const string JobOrderTab = "_tabCustomerProfileJobOrders";

            public const string AddNewJobSiteBtn = "_btnAddJobSite";
            public const string JSActive = "_chkActive";
            public const string JobSiteName = "_txtJobSiteDetailsName";
            public const string JSAddress1 = "_txtJobSiteDetailsAddressLine1";
            public const string JSAddress2 = "_txtJobSiteDetailsAddressLine2";
            public const string JSCountry = "_cmbJobSiteDetailsCountry";
            public const string JSState = "_cmbJobSiteDetailsState";
            public const string JSZip = "_cmbJobSiteDetailsZipCode";
            public const string JSCity = "_cmbJobSiteDetailsCity";
            public const string JSCounty = "_cmbJobSiteDetailsCounty";

            public const string UltraTab = "ultraTabControl1";

            public const string ReportToContactName = "_txtReportToContactName";
            public const string ReportToContactTitle = "_txtReportToContactTitle";
            public const string ReportDepartment = "_txtReportToContactDepartmentSite";
            public const string ReportEmail = "_txtReportToContactEmail";

            public const string JSSave = "_save";
            public const string JSCancel = "_cancel";

            public const string JSSummary = "_grdJobSiteSummary";
            public const string JSStates = "_cmbStates";
            public const string JSCities = "_cmbCities";

            public const string WorkerGrid = "_grdWorkerSummary";

            #endregion

            #region Invoices Controls

            public const string InvTab = "_tabInvoicesPayments";
            public const string InvGrid = "_grdInvoices";
            public const string InvoiceNumber = "_lblInvoiceNumber";
            public const string BalanceAmount = "_txtBalanceAmount";
            public const string JobOrderDetail = "_lnkJobOrderNumber";
            public const string ApplyCC = "_btnCreditCardPayment";
            public const string ReprintOriginal = "_btnReprintOriginal";
            public const string PrintAdjusted = "_btnPrintAdjustedInvoice";

            public const string SinceDate = "_ccSinceDate";
            public const string HistoryFilter = "_btnHistoryFilterGo";
            public const string PaymentHistoryGrid ="_grdPaymentHistory";

            public const string PaymentType ="_cboViewPaymentType";
            public const string GetHistoryFor = "_cbViewHistoryFor";
            public const string JobOrderNumber = "lnkJobOrderNumber";

            public const string ExceptionOk = "ButtonAccept";

            #endregion

            #region Management Controls

            public const string ChangeStatus = "_btnChangeStatus";
            public const string DDChangeStatus = "_cmbCustomerStatus";
            public const string ReasonForStatusChange = "_cmbReasonForStatusChange";
            public const string ReasonForChangeTxt = "_txtReasonForChange";
            public const string AccountStatus ="_lblAccountStatus";

            public const string CustomerLinking ="_btnShowCustomerLinking";
            public const string ParentNumber = "_txtParentCustomerNumber";
            public const string ValidateParent = "_btnValidateParent";

            public const string CreditCheck = "_chkUseParentCredit";
            public const string ParentQuote = "_chkUseParentQuoteProfile";
            public const string ParentCollections = "_chkUseParentCollections";
            public const string OrderingRules = "_chkUseParentOrderProfile";
            public const string InvoiceAndBilling = "_chkUseParentInvoicingBilling";
            public const string SaveCustomerLinking = "_btnSaveCustomerLinking";

            public const string ChainLink = "pboxParentIcon";

            #endregion

            
        }

        public static void SelectSingleOrganization()
        {
            //
            var selectOrgWindow = GetSelectOrgWindowProperties();
            Actions.SetCheckBox(selectOrgWindow, CProfileControls.JobOrderOwner, "False");
            Actions.SetCheckBox(selectOrgWindow, CProfileControls.OrganizationOwner, "True");
            //var invOrg = Actions.GetWindowChild(selectOrgWindow, CProfileControls.InvoicingOrg);
            //Actions.SetCheckBox((WinCheckBox) invOrg, "False");
        }

        public static void EnterJobOrderFindQuoteData(DataRow data)
        {
            var getJobOrderWindow = JobOrderWindow.JobOrderWindow.GetNewJobOrderWindowProperties();

            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            {

                Factory.SetMaskedText(getJobOrderWindow, "txtJobSiteZip", data.ItemArray[9].ToString());

            }

            //ClickOnButton("GO");
            MouseActions.ClickButton(getJobOrderWindow, "btnGo");
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
    }

    public class CCTabConstants
    {
        public const string Summary = "Summary";
        public const string ProfileDetails = "Profile Details";
        public const string CreditInfo = "Credit Info";
        public const string Quotes = "Quotes";
        public const string JobOrders = "Job Orders";
        public const string Invoices = "Invoices && Payments";
        public const string Notes = "Notes";
        public const string Servicing = "Servicing";

        public const string Customer = "Customer";
        public const string Contacts = "Contacts";
        public const string Management = "Management";
        public const string Quoting = "Quoting";
        public const string Billing = "Billing";
        public const string Collections = "Collections";

        public const string CreditDetails = "Credit Details";
        public const string CreditLimit = "Credit Limit";

        public const string CustomerServiceManagement = "Customer Service Management";
        public const string CommonRequirements = "Common Requirements";
        public const string Documents = "Documents";

        #region Quoting Rules Controls

        public const string ManageQuotes = "optManageQuotes";
        public const string PredefinedContact = "optPredefinedContact";
        public const string RateAgreementWindow = "optRateAgreement";
        public const string OverRideBranchThreshold = "_uceOverrideBranchThresholds";
        public const string Save = "_btnSave";
        public const string Cancel = "_btnCancel";

        #endregion

        #region Quote Profile Tabs

        public const string QuoteDetail = "Quote Detail";
        public const string QuoteRates = "Rates";
        public const string DeliverApproveQuote = "Deliver / Approve Quote";
        public const string QuoteHistory = "Quote History";

        #endregion

        #region Job Order Controls

        public const string JobOrdersTab = "Job Orders";
        public const string JobSitesTab = "Job Sites";
        public const string Workers = "Workers";

        public const string JobSiteDetails = "Job Site Details";
        public const string JobSiteContact = "Job Site Contact";
        public const string ReportToContact = "Report To Contact";

        #endregion

        #region Invoices Controls

        public const string InvoicesTab = "Invoices";
        public const string PaymentsTab = "Payments && Credits";
        public const string InvoiceSummary = "Invoice Summary";
        public const string CollectionActivity = "Collection Activity";
        public const string TransactionHistory = "Transaction History";
        public const string JobOrderDetails = "Job Order Detail";
        public const string DispatchLineItemDetail = "Dispatch / Line Item Detail";

        #endregion
    }
}