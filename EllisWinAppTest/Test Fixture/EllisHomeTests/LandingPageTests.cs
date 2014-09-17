using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.EllisHomeTests
{
    [CodedUITest]
    public class LandingPageTests : AppContext
    {
        public void Initialize(int retries = 5)
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();

            if (!App.WaitForControlReady(6000) && retries >= 0)
                Initialize(retries - 1);
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("LandingPage"), TestCategory("Positive")]
        public void VerifyJobOrderSchedulesDisplayed()
        {
            try
            {
                Initialize();

                LandingPage.SelectFromToolbar("Job Orders");
                Factory.AssertIsTrue(LandingPage.VerifyJobOrderSchedulesDisplayed(), "Job Order not displayed");
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("LandingPage"), TestCategory("Positive")]
        public void DispatchMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Dispatch");
            Factory.AssertIsTrue(LandingPage.VerifyDispatchDisplayed(), "Dispatch Menu Items not displayed");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("LandingPage"), TestCategory("Positive")]
        public void LandingPagesMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Workers");
            Factory.AssertIsTrue(LandingPage.VerifyWorkersDisplayed(), "Active Workers not displayed");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("LandingPage"), TestCategory("Positive")]
        public void CustomersMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("Customers");
            Factory.AssertIsTrue(LandingPage.VerifyCustomersDisplayed(), "Customer Quotes not displayed");
            Cleanup();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("LandingPage"), TestCategory("Positive")]
        public void ARMenuItem()
        {
            Initialize();

            LandingPage.SelectFromToolbar("AR");
            Factory.AssertIsTrue(LandingPage.VerifyArDisplayed(), "Customer Collections not displayed");
            Cleanup();
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}