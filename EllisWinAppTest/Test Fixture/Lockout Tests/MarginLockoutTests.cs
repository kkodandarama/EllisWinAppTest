using System.Collections.Generic;
using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EllisWinAppTest.Windows.SearchWindow;

namespace EllisWinAppTest.CustomerTests
{
    [CodedUITest]
    public class MarginLockoutTests : AppContext
    {
        public void Initialize()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsCSRUser();
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("AdvancedSearch"), TestCategory("Positive")]
        public void LockoutNotificationTest()
        {
            Initialize();
            SearchWindow.SelectSearchElements("Test", "Customer", SearchWindow.SearchTypeConstants.Simple);
            if(CustomerProfileWindow.SelectFirstCustomerFromTable())
            {
                CustomerProfileWindow.ClickCreateNewQuoteButton();
                CustomerProfileWindow.CloseWarningWindow();
                    CustomerProfileWindow.CreateNewQuoteWithDefaultData();

                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "Customer profile window is not displayed");
                CustomerProfileWindow.CloseCustomerProfileWindow();
            }
        }

    }
}
