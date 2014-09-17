using System.Collections.Generic;
using System.Data;
using System.Linq;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.CustomerWindow;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.CustomerTests
{
    [CodedUITest]
    public class CreateCustomerTest : AppContext
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

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateNewCustomerWithoutFedIdTest()
        {
            try
            {
                var datarows = Initialize();
                foreach (var datarow in datarows)
                {
                    CreateCustomerWindow.EnterCustomerData(datarow, null, null);
                    CreateCustomerWindow.ClickSave();
                    Playback.Wait(3000);
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "Customer Profile not displayed for new customer without FED ID");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateNewCustomerWithFedIdTest()
        {
            try
            {
                var datarows = Initialize();
                foreach (var datarow in datarows)
                {
                    CreateCustomerWindow.EnterCustomerData(datarow, "Fed", null);
                    CreateCustomerWindow.ClickSave();
                    CreateCustomerWindow.HandleExistingFEDCustomer();
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "Customer Profile not displayed for new customer with FED ID");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateCODCustomerWithoutFedIdTest()
        {
            try
            {
                var datarows = Initialize();
                foreach (var datarow in datarows.Take(3))
                {
                    CreateCustomerWindow.EnterCustomerData(datarow, null, "COD");
                    CreateCustomerWindow.ClickSave();
                    Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "COD Customer Profile not displayed for new COD customer without FED ID");
                }
            }
            finally
            {
                Cleanup();
            }
        }

        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void CreateCODCustomerWithFedIdTest()
        {
            var datarows = Initialize();
            foreach (var datarow in datarows)
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", "COD");
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();

                //TODO: Change Assert state for demo
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "COD Customer Profile not displayed for new COD customer with FED ID");
            }
            Cleanup();
        }


        [TestMethod]
        [TestCategory("UAT"), TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAPSCreateCustomerTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAPSUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.DiffCredsCreateCustomer);

            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("NAPS")))
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", null);
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "NAPS Customer Profile not displayed for new customer with FED ID");
            }
            Cleanup();
        }


        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void NAAMCreateCustomerTest()
        {
            WindowsActions.KillEllisProcesses();
            App = EllisHome.LaunchEllisAsNAAMUser();
            CreateCustomerWindow.ClickOnCreateCustomer();
            var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.DiffCredsCreateCustomer);

            foreach (var datarow in datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("NAAM")))
            {
                CreateCustomerWindow.EnterCustomerData(datarow, "Fed", null);
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();
                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                    "NAPS Customer Profile not displayed for new customer with FED ID");
            }
            Cleanup();
        }

        [Ignore]
        [TestMethod]
        [TestCategory("Regression"), TestCategory("Customer"), TestCategory("Positive")]
        public void ECUSTCreateCustomerTest()
        {
            try
            {
                WindowsActions.KillEllisProcesses();
                App = EllisHome.LaunchEllisAsCSRUser();
                CreateCustomerWindow.ClickOnCreateCustomer();
                var datarows = ExcelReader.ImportSpreadsheet(ExcelFileNames.DiffCredsCreateCustomer);
                var dataRow = datarows.Where(datarow => datarow.ItemArray[3].ToString().Contains("ECUST")).First();

                CreateCustomerWindow.EnterCustomerData(dataRow, "Fed", null);
                CreateCustomerWindow.ClickSave();
                CreateCustomerWindow.HandleExistingFEDCustomer();

                Factory.AssertIsTrue(CustomerProfileWindow.VerifyCustomerProfileWindowDisplayed(),
                        "NAPS Customer Profile not displayed for new customer with FED ID");
            }
            finally
            {
                Cleanup();
            }               
            
        }

        public void Cleanup()
        {
            EllisHome.ClickOnFileExit();
        }
    }
}