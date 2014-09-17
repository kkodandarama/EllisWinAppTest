using System.IO.Packaging;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.AccountsReceivableWindow;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using EllisWinAppTest.Windows.SearchWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.AccountReceivableTests
{
    [CodedUITest]
    public class AccountReceivableTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
            //App = EllisHome.LaunchEllisAsDiffUserFromDesktop();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void VerifyARLandingPageDefaults()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            Factory.AssertIsTrue(ARWindow.VerifyInvoices("All"), "UnPaid invoices are not displayed as All");
            Factory.AssertIsTrue(ARWindow.VerifyMyOrg("Is Collecting"), "My Org is not equal to Is Collecting on landing page");
            Factory.AssertIsTrue(ARWindow.VerifyOverDueDisplayed(), "Over dues are not displayed");
            //Factory.AssertIsTrue(ARWindow.VerifyCustomerProfileWindowDisplayedWhenCustomerNumberClicked(),
            //    "Customer profile page is not displayed when customer on landing page is clicked");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void VerifyARRescheduleNewNoteTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            if (RightClick.Reschedule())
            {
                Factory.AssertIsTrue(ARWindow.VerifyCustomerCollectionsWindowDisplayed(),
                    "Customer collections window is not displayed");

                ARWindow.AddNewNoteToUnpaidInvoice();
                Factory.AssertIsTrue(ARWindow.VerifyNewNoteAdded(), "New note is not added to the unpaid invoice");
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void VerifyARCancelCallbackNewNoteTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            if (RightClick.CallBack())
            {
                Factory.AssertIsTrue(ARWindow.AddNoteWindowDisplayed(), "Add Note window is not displayed");

                ARWindow.AddNewNoteToCancelCallback();
                Factory.AssertIsTrue(ARWindow.VerifyNewNoteAdded(), "New note is not added to the cancel callback");
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void VerifyARCustomerProfileDefaultsTest()
        {
            Initialize();
            LandingPage.SelectFromToolbar("AR");

            ARWindow.SelectCustomerCollectionFromLandingPage();
            CustomerProfileWindow.SelectTab(CCTabConstants.Invoices);
            Factory.AssertIsTrue(CustomerProfileWindow.VerifyNoteTypeOptionsDisplayed(), "Note type options are not displayed.");

            CustomerProfileWindow.SelectTab(CCTabConstants.Notes);
            ARWindow.ClickAddNoteButton();

            Factory.AssertIsTrue(ARWindow.VerifyBalanceDueAmountDisplayed(), "Balance amount is not displayed");
            Factory.AssertIsTrue(ARWindow.VerifyProfileWindowDisplayedWhenCusNameLinkClicked(),
                "Customer profile window is not displayed when customer name link is clicked");
            CustomerProfileWindow.CloseCustomerProfileWindow();

            //Factory.AssertIsTrue(ARWindow.VerifyProfileWindowDisplayedWhenCusNumberLinkClicked(),
            //    "Customer profile window is not displayed when customer number link is clicked");
            //CustomerProfileWindow.CloseCustomerProfileWindow();
            TitlebarActions.ClickClose((WinWindow)ARWindow.GetCustomerCollectionsWindowProperties());

            LandingPage.SelectEODReviewFromNavigationExplorer();
            ARWindow.SearchForAllEODReviews();

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void LockboxInvoiceSearchTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();

            //LandingPage.ClickOnCalendarButton(LandingPage.LandingPageControls.Advanced);
            //LandingPage.EnterDate(LandingPage.LandingPageControls.AdvancedFromDate, "11/16/2009");
            //LandingPage.ClickDateTextbox(LandingPage.LandingPageControls.AdvancedToDate);
            //Playback.Wait(2000);

            LandingPage.SelectDateRange();

            ARWindow.SelectFirstCustomerInvoiceFromTable();
            ARWindow.SelectFirstRemittenceFromTable();
            ARWindow.ClickOpeninNewWindowButton();

            Factory.AssertIsTrue(ARWindow.VerifyPaymentInvoiceWindowDisplayed(),
                "Payment profile invoice window is not displayed when clicked on Open In New Window");
            Factory.AssertIsTrue(ARWindow.VerifyRemainingAmountDisplayedOnWindow(),
                "Remaining Amount is not displayed on window");

            ARWindow.ClosePaymentInvoiceWindow();
            //TitlebarActions.ClickClose((WinWindow) ARWindow.GetPaymentLockboxWindowProperties());
            //TitlebarActions.ClickClose((WinWindow) ARWindow.GetPaymentProfileWindowProperties());
            Playback.Wait(1000);
            Actions.SendAltF4();
            Playback.Wait(1000);
            Actions.SendAltF4();

            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARApplyTransactionTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();
            LandingPage.SelectFromToolbar("AR");
            Playback.Wait(2000);
            LandingPage.SelectCustomerInvoicesFromNavigationExplorer();

            Factory.AssertIsTrue(ARWindow.VerifyCustomerInvGridDisplayed(), "Customer Invoice Grid is not displayed");
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARDefaultToOwnCostDMUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsDMUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Factory.AssertIsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("1926 - NW Empire"),
                "Default district is not equal to 1926 - NW Empire");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARDefaultToAllCorporateARRUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARRUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Factory.AssertIsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARDefaultToAllCorporateAVPUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsAVPUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "Invoices", SearchWindow.SearchTypeConstants.Advanced);

            Factory.AssertIsTrue(ARAdvancedSearchWindow.VerifyInvoicingOrganizationIsNull(),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARDefaultToAllCorporateNABSUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNABSUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Factory.AssertIsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AR"), TestCategory("Positive")]
        public void ARDefaultToAllCorporateARMUserTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsARMUser();

            LandingPage.SelectFromToolbar("AR");
            SearchWindow.SelectSearchElements(null, "ITransactions", SearchWindow.SearchTypeConstants.Advanced);

            Factory.AssertIsTrue(ARAdvancedSearchWindow.VerifyDefaultDistrictSelected("All"),
                "Default district is not equal to All");

            Cleanup();
        }

        private static void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}