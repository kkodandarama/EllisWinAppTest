using System.Collections.Generic;
using System;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.JobOrderWindow;
using EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace EllisWinAppTest.CustomerTests
{
    [CodedUITest]
    public class CustomerProfileTest : AppContext
    {
        public IEnumerable<DataRow> Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            return datarows;
        }

        public void Initialize(int retries = 5)
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            if (!App.WaitForControlReady(6000) && retries >= 0)
                Initialize(retries - 1);
            }


        //public void VerifyCustomerProfileDefaults()
        //{
        //    var datarows = Initialize();
        //    foreach (var datarow in datarows)
        //    {
        //        CreateCustomerWindow.EnterCustomerData(datarow, null, null);
        //        CreateCustomerWindow.ClickSave();

        //        CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
        //        Factory.AssertIsTrue(CustomerProfileWindow.VerifyProfileDefaults(),
        //            "Profile displayed is not that of the customer selected");
        //    }
        //}

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void VerifyNABSCustomerProfileDefaults()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyProfileDefaults(),
                    "Profile displayed is not that of the customer selected");
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateNewCustomerContactTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

                var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
                foreach (var contactData in datarows2)
                {
                    CustomerProfileWindow.ClickAddContactButton();
                    CustomerProfileWindow.EnterContactDetails(contactData);
                    CustomerProfileWindow.ClickSave();
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewContactAdded(Globals.CustomerContact),
                        "New Contact information is not added to the customer profile");
                    break;
                }
                break;
            }
            TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetCustomerProfileWindowProperties());
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateNewNABSCustomerContactTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomer);
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                CreateCustomerWindow.ClickSave();

                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

                var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
                foreach (var contactData in datarows2)
                {
                    CustomerProfileWindow.ClickAddContactButton();
                    CustomerProfileWindow.EnterContactDetails(contactData);
                    CustomerProfileWindow.ClickSave();
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewContactAdded(Globals.CustomerContact),
                        "New Contact information is not added to the customer profile");
                    break;
                }
                break;
            }
            TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetCustomerProfileWindowProperties());
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void EditCustomerContactsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //LandingPage.SelectFromToolbar("Customers");
            //LandingPage.SelectDateRange();
            //if(CustomerProfileWindow.SelectFirstCustomerFromTable())
            //{

            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

            var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var contactData in datarows2)
            {
                CustomerProfileWindow.SelectStatusFilter("All");
                CustomerProfileWindow.SelectFirstCustomerContactFromTable();
                CustomerProfileWindow.EnterContactDetails(contactData);
                CustomerProfileWindow.ClickSave();
                break;
            }

            CustomerProfileWindow.SelectStatusFilter("Active");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewContactAdded(Globals.CustomerContact),
                "Customer contact was not updated");
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void EditActiveContactsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Contacts);

            CustomerProfileWindow.SelectStatusFilter("Active");
            CustomerProfileWindow.SelectFirstCustomerContactFromTable();
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyActiveCheckDisabled(), "Active check is not disabled");
            TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetCustomerContactInfoWindowProperties());
            //TitlebarActions.ClickClose((WinWindow) CustomerProfileWindow.GetCustomerProfileWindowProperties());
            //}

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void AddWorkerCompToCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void AddWorkerCompToNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void AddWorkerCompToTACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.AddWorkersCompToProfile();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                "Workers comp not added successfully");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void AddWorkerCompToRACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsRAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            SearchWindow.SelectFirstCustomerNameFromResults();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            CustomerProfileWindow.ClickWorkersCompButton();
            CustomerProfileWindow.ClickAssociatedAddCode();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyAssociatedCodeTypeSetToPrimary(), "Associated code type is not set to Primary");

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyModifierSetToOne(),
                "Modifier is not set to one");

            CustomerProfileWindow.ClickAddModifierWindow();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyAddModifierWindowDisplayed(), "Add modifier window is not displayed");

            CustomerProfileWindow.CloseAddModifierWindow();
            TitlebarActions.ClickClose((WinWindow)CustomerProfileWindow.GetEditCompCodeModifierWindowProperties());
            CustomerProfileWindow.ClickAssociatedCodeCancel();

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void AddWorkerOrganizationToCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Management);

            if (CustomerProfileWindow.VerifySelectOrganizationButtonEnabled())
            {
                CustomerProfileWindow.ClickSelectOrganizationButton();
                CustomerProfileWindow.SelectOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyOrganizationDetailsUpdated(),
                    "Organization Details were not added successfully");
            }
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void VerifyQuotingRulesForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Quoting);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.ManageQuotes),
                "Manage Quotes option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.PredefinedContact),
                "Predefined Contract option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.RateAgreementWindow),
                "Rate Agreement option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.OverRideBranchThreshold),
                "Over Ride branch threshold option is enabled.");

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Save),
                "Save button is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Cancel),
                "Cancel button is enabled.");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void VerifyQuotingRulesForNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Quoting);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.ManageQuotes),
                "Manage Quotes option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.PredefinedContact),
                "Predefined Contract option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.RateAgreementWindow),
                "Rate Agreement option is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.OverRideBranchThreshold),
                "Over Ride branch threshold option is enabled.");

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Save),
                "Save button is enabled.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingRulesDisabled(CCTabConstants.Cancel),
                "Cancel button is enabled.");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ValidateAndAddNewNoteForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            var notes = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var note in notes)
            {
                CustomerProfileWindow.ClickAddNewNoteButton("General");
                CustomerProfileWindow.AddNewNote(note.ItemArray[3].ToString());
                CustomerProfileWindow.ClickOnSaveAndCloseButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewNoteDisplayedInGrid(note.ItemArray[3].ToString()),
                    "New Note is not added to the customer profile");
            }
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ValidateAndAddNewNoteForNABSCustomerProfileTest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            var notes = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateCustomerContact);
            foreach (var note in notes)
            {
                    CustomerProfileWindow.ClickAddNewNoteButton(note.ItemArray[12].ToString());
                    CustomerProfileWindow.AddNewNote(note.ItemArray[6].ToString());
                CustomerProfileWindow.ClickOnSaveAndCloseButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewNoteDisplayedInGrid(note.ItemArray[4].ToString()),
                    "New Note is not added to the customer profile");
            }
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ValidateNotesForGoAndClearButtonInCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Notes);

            CustomerProfileWindow.EnterDate("From", "01/01/2013");
            CustomerProfileWindow.EnterDate("To", "02/02/2013");

            CustomerProfileWindow.ClickGoButton();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyEmptyGridDisplayed(), "GO button did not function as expected");

            CustomerProfileWindow.ClickClearButton();
            Playback.Wait(200);
            CustomerProfileWindow.ClickGoButton();
            //Factory.AssertIsFalse(CustomerProfileWindow.VerifyEmptyGridDisplayed(), "Clear button did not function as expected");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ValidateCreditHistoryLimitAndDetailsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.CreditInfo);
            CustomerProfileWindow.SelectCreditInfoSubTab(CCTabConstants.CreditLimit);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditStatus("Acceptable"),
                "Credit Status is not displayed as Acceptable. Displayed Status is - " + Globals.Temp);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditTerms(),
                "Credit terms are not equal to either Credit Or COD. Displayed credit term is - " + Globals.Temp);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCurrentCreditLimitDisplayed(),
                "Current Credit Limit is not displayed");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditLimitNotesDisplayed(),
                "Credit Limit Notes grid is not displayed");

            CustomerProfileWindow.SelectCreditInfoSubTab(CCTabConstants.CreditDetails);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditStatusDropDown("Acceptable"),
                "Credit Status is not displayed as Acceptable. Displayed Status is - " + Globals.Temp);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditTermsDropDown(),
                "Credit terms are not equal to either Credit Or COD. Displayed credit term is - " + Globals.Temp);
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Credit Application"),
                "Credit On Application File is not checked");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyDateSignedOnDisplayed(), "Date signed on was not displayed");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyBusinessLicensesDisplayed(),
                "Business Licences were not displayed");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingInvoicingManagementForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.CheckSingleOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                "Billing coordinator displayed is not the selected one.");

            CustomerProfileWindow.ClickCancel();
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingInvoicingManagementForNABSCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.CheckSingleOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                "Billing coordinator displayed is not the selected one.");

            CustomerProfileWindow.ClickCancel();
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingInvoicingManagementForTACustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            if (SearchWindow.SelectFirstCustomerNameFromResults())
            {
                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

                CustomerProfileWindow.ClickSelectOrganizationButton();
                CustomerProfileWindow.UncheckJobOrderOwner();
                CustomerProfileWindow.CheckSingleOrganization();
                CustomerProfileWindow.SelectAllOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyBillingCoordinator(),
                    "Billing coordinator displayed is not the selected one.");

                CustomerProfileWindow.ClickCancel();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingInvoiceDeliveryOptionsForCustomerProfileTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("EBilling"),
                "Customer have opted not to use e-billing");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingInvoicingPreferencesTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.CheckConsolidatedInvoicing();
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Consolidated Invoicing"),
                "Consolidated invoicing option is not checked");

            CustomerProfileWindow.CheckOvertimeBilling();
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Overtime Billing"),
                " Over time billing option is not checked");
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void BillingTaxExemptionsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

            CustomerProfileWindow.ClickViewTaxExceptions();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyExceptionStatusSetToAll(),
                "Exception status is not set to ALL. It is currently set to - " + Globals.Temp);
            CustomerProfileWindow.ClickExcemptionWindowCancel();
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void TABillingTaxExemptionsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Simple);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            if (SearchWindow.SelectFirstCustomerNameFromResults())
            {
                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Billing);

                CustomerProfileWindow.ClickViewTaxExceptions();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyExceptionStatusSetToAll(),
                    "Exception status is not set to ALL. It is currently set to - " + Globals.Temp);
                CustomerProfileWindow.ClickExcemptionWindowCancel();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void SelectCollectionsOrganizationTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (SearchWindow.SelectFirstCustomerNameFromResults())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyAgingThesholdsSectionDisplayed(),
                "Aging thresholds section not displayed as per design");

            CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
            CustomerProfileWindow.UncheckInvoicingOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPrimaryOrgUnitDisplayed(),
                "Primary Organization is not displayed as selected");
            CustomerProfileWindow.ClickCancel();
            //}
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void SelectNABSCollectionsOrganizationTest()
        {
            try
            {
                WindowsActions.KillEllisProcesses();
                App = EllisHome.LaunchEllisAsNABSUser();

                LandingPage.VerifyTestCustomerPresent();
                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

                CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
                CustomerProfileWindow.UncheckInvoicingOrganization();
                CustomerProfileWindow.SelectAllOrganizationDetails();
                CustomerProfileWindow.ClickSelectButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyPrimaryOrgUnitDisplayed(),
                    "Primary Organization is not displayed as selected");
                CustomerProfileWindow.ClickCancel();
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void SelectCollectionsCOITest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            CustomerProfileWindow.ClickPrintCOIButton();
            Playback.Wait(3000);
            CustomerProfileWindow.ClickPrintButton();
            Playback.Wait(3000);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPrintPreviewWindowDisplayed(),
                "Print Preview window is not displayed");
            CustomerProfileWindow.ClosePPWindow();
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ValidateCollectionsFinanceChargesTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Collections);

            Factory.AssertIsTrue(CustomerProfileWindow.SelectFinanceCharges("Yes"), "Radio buttons cannot be checked");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ServicingTabCustomerServiceManagementTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            //SearchWindow.SelectSearchElements("Test_Customer", "Customer", SearchWindow.SearchTypeConstants.Simple);
            //if (CustomerProfileWindow.SelectFirstCustomerFromTable())
            //{
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                "Account manager checkbox is enabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                "Assigned to branch is disabled");
            Factory.AssertIsTrue(CustomerProfileWindow.SelectServiceCoordinator("Test EllisCSR"),
                "Service Coordinator is not enabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                "Authorized to order checkbox is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                "Alert user for servicing rules is disabled");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                "User is permitted to access e-commerce site");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                "Branch permitted check is enabled");

            CustomerProfileWindow.ClickCancel();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            // }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NABSServicingTabCustomerServiceManagementTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                "Account manager checkbox is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                "Assigned to branch is disabled");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyServiceCoordinatorEnabled(),
                "Service Coordinator is not enabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                "Authorized to order checkbox is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                "Alert user for servicing rules is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                "User is not permitted to access e-commerce site");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                "Branch permitted check is enabled");

            CustomerProfileWindow.ClickCancel();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            //CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ServicingTabCommonRequirementsTest()
        {
            try
            {
                Initialize(5);
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Purchase Order Required"),
                "Purchase order required checkbox is diabled");
            Factory.AssertIsFalse(CustomerProfileWindow.SetPOFormat("pdf"), "Not able to set Po format text");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Drug Test"),
                "Drug test required is not enabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Background Check"),
                "Background check checkbox is disabled");

            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void RAServicingTabCommonRequirementsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsRAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            if (SearchWindow.SelectFirstCustomerNameFromResults())
            {
                CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

                CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

                Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Bahavioral Survey"),
                    "Bahavioral Survey Required checkbox is diabled");

                CustomerProfileWindow.CloseCustomerProfileWindow();
                TitlebarActions.ClickClose(SearchWindow.GetSearchResultsWindowProperties());
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ServicingTabDocumentsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectFromProfileDetailsTab(CCTabConstants.Servicing);

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.Documents);

            CustomerProfileWindow.ClickAddFileButton();
            CustomerProfileWindow.FindCustomerDocumentOnFile();
            CustomerProfileWindow.ClickCloseOnCustomerDOF();
            CustomerProfileWindow.ClickCloseOnDOF();

            CustomerProfileWindow.CloseCustomerProfileWindow();
            CustomerProfileWindow.CloseCustomerSearchResultsWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            if (CustomerProfileWindow.SelectFirstQuoteFromResults())
            {
                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(), "Quote status is not displayed or is null");

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                    "Pricing Matrix Grid is not displayed or is null");

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                    "Quote History Grid is not displayed or is null");

                CustomerProfileWindow.CloseQuoteProfileWindow();
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerCreateNewQuoteTest()
        {
            try
            {
                Initialize(5);
            LandingPage.VerifyTestCustomerPresent();
            CustomerProfileWindow.ClickCreateNewQuoteButton();
            CustomerProfileWindow.CloseWarningWindow();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
                var dataRow = datarows.Where(x => x.ItemArray[3].ToString().Contains("Test_Requester")).FirstOrDefault();
                Factory.AssertIsFalse(dataRow == null, "Data could not be found for CustomerCreateNewQuoteTest().");
                CustomerProfileWindow.CreateNewQuote(dataRow);

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "Customer profile window is not displayed");
            CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerCopyQuotesDefaultsTest()
        {
            Initialize(5);

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            //Factory.AssertIsTrue(CustomerProfileWindow.SelectFirstQuoteFromResults(), "No Quote Results displayed in the grid");
            if (CustomerProfileWindow.SelectFirstQuoteFromResults())
            {

                CustomerProfileWindow.ClickCopyQuote();

                Playback.Wait(2000);
                CustomerProfileWindow.CloseWarningWindow();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed(), "");

                if (CustomerProfileWindow.ClickChangeQuoteCustomerName())
                {
                    CustomerProfileWindow.SelectNewCustomerNameAfterSearch("Test");

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyNoQuotingAllowedWindowDisplayed(),
                        "Verify quoteing allowed window not displayed for new customer by name - TEST");
                    CustomerProfileWindow.CloseNoQuotingAllowedWindow();

                    TitlebarActions.ClickClose(GlobalWindows.QuoteProfileWindow);
                    CustomerProfileWindow.CloseUnsavedChangesWindow();
                }
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NABSCustomerQuotesDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(), "Quote status is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                "Pricing Matrix Grid is not displayed or is null");

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

            CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                "Quote History Grid is not displayed or is null");

            CustomerProfileWindow.CloseQuoteProfileWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NABSCustomerCreateNewQuoteTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewQuoteButton();
            CustomerProfileWindow.CloseWarningWindow();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("NABS")))
            {
                if (CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed())
                {
                    CustomerProfileWindow.CreateNewQuote(datarow);

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "Customer profile window is not displayed");
                }
                break;
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NABSCustomerCopyQuotesDefaultsTest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
            CustomerProfileWindow.EnterQuoteSearchCriteria();

            CustomerProfileWindow.SelectFirstQuoteFromResults();
            CustomerProfileWindow.ClickCopyQuote();

            Playback.Wait(2000);
            CustomerProfileWindow.CloseWarningWindow();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed(), "");

            CustomerProfileWindow.ClickChangeQuoteCustomerName();
            CustomerProfileWindow.SelectNewCustomerNameAfterSearch("Test");

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyNoQuotingAllowedWindowDisplayed(),
                "Verify quoteing allowed window not displayed for new customer by name - TEST");
            CustomerProfileWindow.CloseNoQuotingAllowedWindow();

            TitlebarActions.ClickClose(GlobalWindows.QuoteProfileWindow);
            CustomerProfileWindow.CloseUnsavedChangesWindow();
            }
            finally
            {
            Cleanup();
        }

        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void TACustomerQuotesDefaultsTest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
                var customerFound = SearchWindow.SelectFirstCustomerNameFromResults();
                if (customerFound)
            {
                CustomerProfileWindow.SelectTab(CCTabConstants.Quotes);
                CustomerProfileWindow.EnterQuoteSearchCriteria();

                CustomerProfileWindow.SelectFirstQuoteFromResults();
                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteDetail);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteStatusDisplayed(),
                    "Quote status is not displayed or is null");

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteRates);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyPricingMatrixDisplayed(),
                    "Pricing Matrix Grid is not displayed or is null");

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.DeliverApproveQuote);

                CustomerProfileWindow.SelectFromQuoteProfileTab(CCTabConstants.QuoteHistory);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuoteHistoryGridTable(),
                    "Quote History Grid is not displayed or is null");

                CustomerProfileWindow.CloseQuoteProfileWindow();
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
                Factory.AssertIsTrue(customerFound, "Could not find a customer.");
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerVerifyCreateJobOrderWindowDisplayedTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();
            Factory.AssertIsTrue(JobOrderWindow.VerifyCreateJobOrderWindowDisplayed(),
                "Job Order window is not displayed when clicked on Create job Order button");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerJobOrdersWorkflowTest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);

            CustomerProfileWindow.EnterJobOrderSearchCriteria();
            CustomerProfileWindow.SelectFirstJobOrderFromResults();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyBasicJobInfoDisplayed(), "Basic job info is not displayed");

            TitlebarActions.ClickClose((WinWindow)OpenJobOrder.JobOrderProfileWindowProperties());
            CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerJOAddNewJobSiteWorkflowTest()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
            CustomerProfileWindow.SelectJOSubTab(CCTabConstants.JobSitesTab);

            CustomerProfileWindow.ClickAddNewJobSiteButton();

            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateJobSite);
            foreach (var datarow in datarows)
            {
                CustomerProfileWindow.CreateNewJobSite(datarow);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyNewJobSiteCreated(datarow, Globals.Temp),
                    "New Job site is not created");
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            finally
            {
            Cleanup();
        }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void TACustomerJOAddNewJobSiteWorkflowTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsTAUser();

            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.EnterCustomerNameAsSearchData("Test");
            if (SearchWindow.SelectFirstCustomerNameFromResults())
            {
                CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
                CustomerProfileWindow.SelectJOSubTab(CCTabConstants.JobSitesTab);

                CustomerProfileWindow.ClickAddNewJobSiteButton();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyAddNewJobSiteWindowDispllayed(),
                    "Add new job site is not displayed");
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void VerifyCustomerJOWorkersTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.JobOrders);
            CustomerProfileWindow.SelectJOSubTab(CCTabConstants.Workers);

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyWorkersDisplayed(), "workers table is not displayed");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerCreateNewJobOrderTest()
        {
            Initialize(5);

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();
            CustomerProfileWindow.CloseWarningWindow();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);

            // Find Customer Window
            //var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            var dataRow = datarows.Where(x => JobOrderWindow.CreateNewJobOrder(x)).FirstOrDefault();
            Factory.AssertIsFalse(dataRow == null, "Couldn't find a suitable datarow for CustomerCreateNewJobOrderTest().");

                    // Find Quote Tab/Window
                    Playback.Wait(3000);
                    JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
            //JobOrderFindQuoteWindow.ClickOnButton("GO");
            //Playback.Wait(2000);
                    JobOrderWindow.ClickOnContinueBtn();

                    // Enter Basic Job Order Details
                    BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                    BasicJobInformationWindow.ClickOnContinueBtn();

                    // Enter Schedule And Additional Charges Details
                    ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                    ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                    // Enter Order Notes in Schedule And Additional Charges window
                    ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                    // Focus back to Schedule And Additional Charges window
                    ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                    //Enter data in Requirements window
                    RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                    RequirementsWindow.ClickOnButton("Continue >");
                    Playback.Wait(3000);

                    //Enter data in Pre-Qualifying Requirements Window

                    PreQualifyingQuestionsWindow.ClickonSaveButton();
                    PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                    PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                    Playback.Wait(3000);
                    var status = PreQualifyingQuestionsWindow.HandleAlertWindow();
                    Factory.AssertIsTrue(status, "Job order was not created.");

        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NABSCustomerCreateNewJobOrderTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.ClickCreateNewJobOrderButton();
            CustomerProfileWindow.CloseWarningWindow();
            var datarows = EllisHome.Initialize(ExcelFileNames.JobOrder);

            // Find Customer Window
            //var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows.Where(dataRow => dataRow.ItemArray[4].ToString().Contains("NABS")))
            {
                //Console.WriteLine(dataRow.ItemArray[24]);
                JobOrderWindow.EnterJobOrderData(dataRow);
                JobOrderWindow.ClickOnButton("Search");

                Playback.Wait(3000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnButton("Continue");
                //JobOrderWindow.ClickOnContinueBtn();

                // Find Quote Tab/Window
                Playback.Wait(3000);
                JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                //JobOrderFindQuoteWindow.ClickOnButton("GO");

                Playback.Wait(2000);
                Actions.SendText("%C");
                //JobOrderWindow.ClickOnContinueBtn();

                // Enter Basic Job Order Details
                BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                BasicJobInformationWindow.ClickOnContinueBtn();

                // Enter Schedule And Additional Charges Details
                ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                // Enter Order Notes in Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);

                // Focus back to Schedule And Additional Charges window
                ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();

                //Enter data in Requirements window
                RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                RequirementsWindow.ClickOnButton("Continue >");
                Playback.Wait(3000);

                //Enter data in Pre-Qualifying Requirements Window
                PreQualifyingQuestionsWindow.ClickonSaveButton();
                PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                Playback.Wait(3000);
                PreQualifyingQuestionsWindow.HandleAlertWindow();

                Factory.AssertIsTrue(PreQualifyingQuestionsWindow.HandleAlertWindow(), "Job order not saved successfully");
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerInvoicesTabTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoicesTab);

            CustomerProfileWindow.SearchAndSelectFirstInvoiceFromGrid();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerInvoiceNumberDisplayed(Globals.Temp),
                "Invoice number is not displayed on the customer invoice window");

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.TransactionHistory);
            CustomerProfileWindow.SearchForTransactionHistory();

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.DispatchLineItemDetail);
            CustomerProfileWindow.ClickOnJobOrderNumberLink();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyJobOrderWindowDisplayed(), "");

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.InvoiceSummary);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyBalanceDueDisplayed(), "Balance due is not displayed");

            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.JobOrderDetails);
            CustomerProfileWindow.ClickOnJobOrderNumber();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyJobOrderWindowDisplayed(), "Job order window is not displayed");

            try
            {
                TitlebarActions.ClickClose(GlobalWindows.JobOrderProfileWindow);
            }
            catch
            {
                GlobalWindows.JobOrderProfileWindow.SetFocus();
                Actions.SendAltF4();
            }
            

            if (CustomerProfileWindow.ApplyCcPaymentEnabled())
            {
                CustomerProfileWindow.ClickApplyCCPayment();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyApplyCCPaymentWindowDisplayed(),
                    "Apply CC payment window is not displayed.");
                CustomerProfileWindow.CloseApplyCCPaymentWindow();

            }

            CustomerProfileWindow.ClickOnReprintOriginalButton();
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPrintDialogWindowDisplayed(),
                "Print window is not displayed.");
            Actions.SendAltF4();
            //TitlebarActions.ClickClose(GlobalWindows.DialogWindow);

            CustomerProfileWindow.ClickOnPrintAdjustedButton();
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyEllisExceptionWindowDisplayed(),
                "Ellis Exception Window is displayed displayed for Print Adjusted action.");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPrintDialogWindowDisplayed(),
                "Print window is not displayed.");
            Actions.SendAltF4();
            //TitlebarActions.ClickClose(GlobalWindows.DialogWindow);

            CustomerProfileWindow.CloseCustomerInvoiceWindow();
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CustomerInvoicesPaymentsAndCreditsTabTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectFirstCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.PaymentsTab);

            CustomerProfileWindow.SearchForAllInvoices();
            //CustomerProfileWindow.VerifySearchResultsDisplayed();

            if (CustomerProfileWindow.VerifySearchResultsDisplayed())
            {
                Factory.AssertIsTrue(CustomerProfileWindow.VerifySearchResultsDisplayed(), "Search results are not displayed");

                CustomerProfileWindow.ClickClearButton();
                CustomerProfileWindow.ClickHistoryFilter();
                Factory.AssertIsFalse(CustomerProfileWindow.VerifySearchResultsDisplayed(), "Search results are displayed");
            }
            CustomerProfileWindow.CloseCustomerProfileWindow();
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAPSCustomerDefaultsTest()
        {
            try
            {
                WindowsActions.KillEllisProcesses();
                App = EllisHome.LaunchEllisAsNAPSUser();

                LandingPage.SelectFromToolbar("Customers");
                if (CustomerProfileWindow.SelectRandomCustomerFromTable())
                {

                    CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
                    CustomerProfileWindow.SelectCustomerSegmentation("National");
                    //Factory.AssertIsTrue(CustomerProfileWindow.VerifyEllisExceptionWindowDisplayed(), "Ellis exception window is displayed");
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerSegment("National"),
                        "Customer segment changes once save button is clicked");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Management);
                    //CustomerProfileWindow.ClickChangeStatus();
                    CustomerProfileWindow.ChangeStatus("Do Not Service");
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyStatusChanged("Do Not Service"),
                        "Status has not been changed to Do Not Service");

                    CustomerProfileWindow.ClickChangeStatus();
                    CustomerProfileWindow.ChangeStatus("Approved");
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyStatusChanged("Approved"),
                        "Status has not been changed to Approved");

                    CustomerProfileWindow.ClickCustomerLinking();
                    CustomerProfileWindow.EnterParentCustomerNumber();
                    CustomerProfileWindow.ClickValidate();

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyInheritanceOptionsEnabled(),
                        "Inheritance options are not enabled. Please check entered parent customer number");
                    CustomerProfileWindow.ClickSaveCustomerLinking();
                    Thread.Sleep(3000);
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyChainLinkDisplayed(),
                        "Customer linking Icon is not displayed");

                    CustomerProfileWindow.ClickCustomerLinking();
                    CustomerProfileWindow.DelinkParentCustomer();
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyChainLinkDisplayed(),
                        "Customer linking Icon is still displayed");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Quoting);
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingCheckboxesEnabled(),
                        "Quoting checkboxes are not checked");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Collections);
                    CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
                    //This is unchecked the first time.
                    //CustomerProfileWindow.UncheckInvoicingOrganization();
                    CustomerProfileWindow.SelectAllOrganizationDetails();
                    CustomerProfileWindow.ClickSelectButton();

                    CustomerProfileWindow.SelectTab(CCTabConstants.Billing);
                    CustomerProfileWindow.ClickSelectOrganizationButton();
                    //CustomerProfileWindow.UncheckInvoicingOrganization();
                    CustomerProfileWindow.SelectSingleOrganization();
                    CustomerProfileWindow.SelectAllOrganizationDetails();
                    CustomerProfileWindow.ClickSelectButton();
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyInvOrgDisplayed(),
                        "Selected Invoicing Org is Displayed");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
                    CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoicesTab);
                    CustomerProfileWindow.SearchForUnPaidInvoices();

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyUnpaidInvoicesDisplayed(),
                        "Unpaid invoices are not available for the selected customer");

                    if (CustomerProfileWindow.VerifyUnpaidInvoicesDisplayed())
                    {
                        CustomerProfileWindow.ClickApplyCCPayment();
                        var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreditCard);

                        foreach (var datarow in datarows)
                        {
                            CustomerProfileWindow.EnterCreditCardDetails(datarow);
                            CustomerProfileWindow.ClickProcessPayment();
                            CustomerProfileWindow.CloseApplyCCPaymentWindow();
                            break;
                        }
                    }

                    CustomerProfileWindow.SelectTab(CCTabConstants.CreditInfo);
                    CustomerProfileWindow.SelectCreditInfoSubTab(CCTabConstants.CreditLimit);

                    CustomerProfileWindow.ClickUpdateCreditLimit();
                    CustomerProfileWindow.AddNewCreditLimit("50000");
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCurrentCreditLimitDisplayed("$500.00"),
                        "Credit limit displayed is not as per the new credit limit entered");

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditLimitWarnAndLockoutDisplayed(),
                        "Credit limit warn and credit limit lockout is not displayed");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Servicing);
                    CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                        "Account manager checkbox is disabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                        "Assigned to branch is disabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                        "Authorized to order checkbox is disabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                        "Alert user for servicing rules is disabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                        "User is not permitted to access e-commerce site");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                        "Branch permitted check is enabled");

                    CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Purchase Order Required"),
                        "Purchase order required checkbox is diabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.SetPOFormat("pdf"), "Not able to set Po format text");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Drug Test"),
                        "Drug test required is not enabled");
                    Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Background Check"),
                        "Background check checkbox is disabled");

                    CustomerProfileWindow.SelectTab(CCTabConstants.Summary);
                    CustomerProfileWindow.ClickCreateNewQuoteButton();
                    CustomerProfileWindow.CloseWarningWindow();

                    var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
                    foreach (var datarow in datarows2.Where(datarow => datarow.ItemArray[3].ToString().Contains("NAPS"))
                        )
                    {
                        CustomerProfileWindow.CreateNewQuote(datarow);

                        Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                            "Customer profile window is not displayed");

                        CustomerProfileWindow.ClickCreateNewJobOrderButton();
                        CustomerProfileWindow.CloseWarningWindow();

                        break;
                    }



                    var datarows3 = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
                    foreach (var dataRow in datarows3)
                    {
                        if (dataRow.ItemArray[1].Equals("CustomerDefaultsTest"))
                        {
                            JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                            Playback.Wait(2000);
                            //Alt + C ?
                            Actions.SendText("%C");
                            BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                            BasicJobInformationWindow.ClickOnContinueBtn();

                            ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                            ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                            ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);
                            ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();
                            RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                            RequirementsWindow.ClickOnButton("Continue >");
                            Playback.Wait(3000);
                            PreQualifyingQuestionsWindow.ClickonSaveButton();
                            PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                            PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                            Playback.Wait(3000);
                            PreQualifyingQuestionsWindow.HandleAlertWindow();

                            Factory.AssertIsTrue(PreQualifyingQuestionsWindow.HandleAlertWindow(),
                                "Job order not saved successfully");
                        }
                        break;
                    }

                    Cleanup();
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAPSRequestSalesCredit()
        {
            try
            {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAPSUser();
                App.Process.WaitForInputIdle();
            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.SearchByCustomerNumber("3584-2159");

            Globals.CustomerName = "WM SIMI VALLEY LANDFILL #4108";
            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectFirstInvoiceFromGrid();
            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoiceSummary);
            CustomerProfileWindow.ClickRequestCredit();
            CustomerProfileWindow.EnterCreditRequestData();
            TitlebarActions.ClickClose(CustomerProfileWindow.GetCustomerInvoiceWindowProperties());
            CustomerProfileWindow.SelectFirstInvoiceFromGrid();

            CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoiceSummary);
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyPendingCreditRequestAmount("10.00"), "");
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAAMCustomerDefaultsTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAAMUser();

            LandingPage.SelectFromToolbar("Customers");
            CustomerProfileWindow.SelectRandomCustomerFromTable();

            CustomerProfileWindow.SelectTab(CCTabConstants.ProfileDetails);
            CustomerProfileWindow.SelectCustomerSegmentation("National");
            //Factory.AssertIsTrue(CustomerProfileWindow.VerifyEllisExceptionWindowDisplayed(), "Ellis exception window is displayed");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerSegment("National"), "Customer segment changes once save button is clicked");

            CustomerProfileWindow.SelectTab(CCTabConstants.Management);
            CustomerProfileWindow.ClickChangeStatus();
            CustomerProfileWindow.ChangeStatus("Do Not Service");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyStatusChanged("Do Not Service"), "Status has not been changed to Do Not Service");

            CustomerProfileWindow.ClickChangeStatus();
            CustomerProfileWindow.ChangeStatus("Approved");
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyStatusChanged("Approved"), "Status has not been changed to Approved");

            CustomerProfileWindow.ClickCustomerLinking();
            CustomerProfileWindow.EnterParentCustomerNumber();
            CustomerProfileWindow.ClickValidate();

            Factory.AssertIsTrue(CustomerProfileWindow.VerifyInheritanceOptionsEnabled(), "Inheritance options are not enabled. Please check entered parent customer number");
            CustomerProfileWindow.ClickSaveCustomerLinking();
            Thread.Sleep(3000);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyChainLinkDisplayed(), "Customer linking Icon is not displayed");

            CustomerProfileWindow.ClickCustomerLinking();
            CustomerProfileWindow.DelinkParentCustomer();
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyChainLinkDisplayed(), "Customer linking Icon is still displayed");

            CustomerProfileWindow.SelectTab(CCTabConstants.Quoting);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyQuotingCheckboxesEnabled(), "Quoting checkboxes are not checked");

            CustomerProfileWindow.SelectTab(CCTabConstants.Collections);
            CustomerProfileWindow.ClickInvoicingOrgSelectBtn();
            CustomerProfileWindow.UncheckInvoicingOrganization();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();

            CustomerProfileWindow.SelectTab(CCTabConstants.Billing);
            CustomerProfileWindow.ClickSelectOrganizationButton();
            CustomerProfileWindow.UncheckJobOrderOwner();
            CustomerProfileWindow.SelectAllOrganizationDetails();
            CustomerProfileWindow.ClickSelectButton();
            //Factory.AssertIsTrue(CustomerProfileWindow.VerifyInvOrgDisplayed(), "Selected Invoicing Org is Displayed");

            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            //CustomerProfileWindow.SelectInvSubTab(CCTabConstants.InvoicesTab);
            CustomerProfileWindow.SearchForUnPaidInvoices();

            //Factory.AssertIsTrue(CustomerProfileWindow.VerifyUnpaidInvoicesDisplayed(),
            //    "Unpaid invoices are not available for the selected customer");

            if (CustomerProfileWindow.VerifyUnpaidInvoicesDisplayed())
            {
                CustomerProfileWindow.ClickApplyCCPayment();
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreditCard);

                foreach (var datarow in datarows)
                {
                    CustomerProfileWindow.EnterCreditCardDetails(datarow);
                    CustomerProfileWindow.ClickProcessPayment();
                    CustomerProfileWindow.CloseApplyCCPaymentWindow();
                    break;
                }
            }

            CustomerProfileWindow.SelectTab(CCTabConstants.CreditInfo);
            CustomerProfileWindow.SelectCreditInfoSubTab(CCTabConstants.CreditLimit);

            if (CustomerProfileWindow.ClickUpdateCreditLimit())
            {
                CustomerProfileWindow.AddNewCreditLimit("50000");
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCurrentCreditLimitDisplayed("$500.00"),
                    "Credit limit displayed is not as per the new credit limit entered");

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCreditLimitWarnAndLockoutDisplayed(),
                    "Credit limit warn and credit limit lockout is not displayed");
            }

            CustomerProfileWindow.SelectTab(CCTabConstants.Servicing);
            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CustomerServiceManagement);

            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Account Manager"),
                "Account manager checkbox is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Assigned To Branch"),
                "Assigned to branch is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Authorized to Order"),
                "Authorized to order checkbox is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Alert User"),
                "Alert user for servicing rules is disabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("ECommerce Access"),
                "User is not permitted to access e-commerce site");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Branch Permitted"),
                "Branch permitted check is enabled");

            CustomerProfileWindow.SelectServicingInnerTab(CCTabConstants.CommonRequirements);

            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Purchase Order Required"),
                "Purchase order required checkbox is diabled");
            Factory.AssertIsFalse(CustomerProfileWindow.SetPOFormat("pdf"), "Not able to set Po format text");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Drug Test"),
                "Drug test required is not enabled");
            Factory.AssertIsFalse(CustomerProfileWindow.VerifyCheckboxDisabled("Background Check"),
                "Background check checkbox is disabled");

            CustomerProfileWindow.SelectTab(CCTabConstants.Summary);
            CustomerProfileWindow.ClickCreateNewQuoteButton();
            CustomerProfileWindow.CloseWarningWindow();

            var datarows2 = ExcelReader.ImportSpreadsheet(ExcelFileNames.CreateQuote);
            foreach (var datarow in datarows2.Where(datarow => datarow.ItemArray[3].ToString().Contains("NAAM")))
            {
                if (CustomerProfileWindow.VerifyCreateNewQuoteWindowDisplayed())
                {

                    CustomerProfileWindow.CreateNewQuote(datarow);

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "Customer profile window is not displayed");
                }
                break;
            }

            CustomerProfileWindow.ClickCreateNewJobOrderButton();
            CustomerProfileWindow.CloseWarningWindow();
            var datarows3 = ExcelReader.ImportSpreadsheet(ExcelFileNames.JobOrder);
            foreach (var dataRow in datarows3)
            {
                JobOrderFindQuoteWindow.EnterJobOrderFindQuoteData(dataRow);
                Playback.Wait(2000);
                if (!PreQualifyingQuestionsWindow.ChooseAlertWindowProperties().Exists)
                {
                Actions.SendText("%C");
                BasicJobInformationWindow.EnterBasicJobInformationWindowData(dataRow);
                BasicJobInformationWindow.ClickOnContinueBtn();

                ScheduleAndAdditionalChargesWindow.EnterDataInScheduleAndAdditionalChargesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnAddNotesBtn();

                ScheduleAndAdditionalChargesWindow.EnterDataInJobOrderNotesWindow(dataRow);
                ScheduleAndAdditionalChargesWindow.ClickOnContinueBtn();
                RequirementsWindow.EnterDatainRequirementsWindow(dataRow);
                RequirementsWindow.ClickOnButton("Continue >");
                Playback.Wait(3000);
                PreQualifyingQuestionsWindow.ClickonSaveButton();
                PreQualifyingQuestionsWindow.HandleChooseLocationWindow();
                PreQualifyingQuestionsWindow.HandleWorkLocationWindow();
                Playback.Wait(3000);
                PreQualifyingQuestionsWindow.HandleAlertWindow();

                    Factory.AssertIsTrue(PreQualifyingQuestionsWindow.HandleAlertWindow(),
                        "Job order not saved successfully");
                    break;
                }
                PreQualifyingQuestionsWindow.HandleAlertWindow();
                break;
            }

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAAMRequestSalesCredit()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAAMUser();
            SearchWindow.SelectSearchElements(null, "Customer", SearchWindow.SearchTypeConstants.Advanced);
            CustomerAdvanceSearchWindow.SearchByCustomerNumber("3584-2159");

            Globals.CustomerName = "WM SIMI VALLEY LANDFILL #4108";
            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            CustomerProfileWindow.SelectFirstInvoiceFromGrid();
            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.InvoiceSummary);
            CustomerProfileWindow.ClickRequestCredit();
            CustomerProfileWindow.EnterCreditRequestData();

            TitlebarActions.ClickClose(CustomerProfileWindow.GetCustomerInvoiceWindowProperties());
            CustomerProfileWindow.SelectFirstInvoiceFromGrid();
            CustomerProfileWindow.SelectCustomerInvoiceSubTab(CCTabConstants.InvoiceSummary);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyPendingCreditRequestAmount("10.00"), "");
        }

        private void Cleanup()
        {
            try
            {
                EllisHome.ClickOnFileExit();
            }
            
            catch 
            {
                //Suppressing all exceptions
            }
        }
    }
}