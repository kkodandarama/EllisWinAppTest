using System.Data;
using System.Linq;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class ARAdvancedSearchWindow : AppContext
    {

        #region Window Properties

        private static UITestControl GetSearchWindowProperties()
        {
            var searchWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search" });
            return searchWindow;
        }

        private static UITestControl GetSearchResultsWindowProperties()
        {
            var searchResultsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Search Results" });
            return searchResultsWindow;
        }

        
        #endregion

        #region Billing Line Item Search Methods

        public static void EnterBillingSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var customerName = Actions.GetWindowChild(searchWindow, BillingSearchConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[3].ToString());

            var customerNumber = Actions.GetWindowChild(searchWindow, BillingSearchConstants.CustomerNumber);
            Actions.SetText(customerNumber, data.ItemArray[4].ToString());

            var dispatchStart = Actions.GetWindowChild(searchWindow, BillingSearchConstants.DispatchStart);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
            DropDownActions.SelectDropdownByText(dispatchStart,data.ItemArray[6].ToString());

            var dispatchEnd = Actions.GetWindowChild(searchWindow, BillingSearchConstants.DispatchEnd);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            DropDownActions.SelectDropdownByText(dispatchEnd, data.ItemArray[7].ToString());

            var district = Actions.GetWindowChild(searchWindow, BillingSearchConstants.District);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            DropDownActions.SelectDropdownByText(district, data.ItemArray[8].ToString());

            var branch = Actions.GetWindowChild(searchWindow, BillingSearchConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            DropDownActions.SelectDropdownByText(branch, data.ItemArray[9].ToString());
        }

        public static bool SelectBillingRecord(string billing)
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Enabled)
            {
                var label = Actions.GetWindowChild(searchResultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var billinglineItem = TableActions.OpenRecordFromTable(searchResultsWindow, "_grdSearchResult",
                    "Customer Number", billing);
                return billinglineItem;
            }
            return false;
        }
        #endregion

        #region Credit Card Search Methods

        public static void EnterCreditCardSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var creditcardNo = Actions.GetWindowChild(searchWindow, CreditCardSearchConstants.CreditCard);
            Actions.SetText(creditcardNo,data.ItemArray[11].ToString());

            var tStatus = Actions.GetWindowChild(searchWindow, CreditCardSearchConstants.Transaction);
            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
            DropDownActions.SelectDropdownByText(tStatus,data.ItemArray[12].ToString());

            var from = Actions.GetWindowChild(searchWindow, CreditCardSearchConstants.ProcessedFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
            DropDownActions.SelectDropdownByText(from, data.ItemArray[13].ToString());

            var to = Actions.GetWindowChild(searchWindow, CreditCardSearchConstants.ProcessedTo);
            if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
            DropDownActions.SelectDropdownByText(to, data.ItemArray[14].ToString());
          
            
        }

        public static bool SelectCreditCardRecord(string ccNumber)
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Exists)

            {
                var label = Actions.GetWindowChild(searchWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var creditCard = TableActions.OpenRecordFromTable(searchWindow, "_grdSearchResult",
                    "Card Number (Last 4)", ccNumber);
                return creditCard;
            }
            return false;
        }

        #endregion

        #region Invoice Search Methods

        public static void EnterInvoiceSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var invoiceNo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceNumber);
            Actions.SetText(invoiceNo,data.ItemArray[15].ToString());

            var billingNo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.BillingNumber);
            Actions.SetText(billingNo, data.ItemArray[16].ToString());

            var ticketNo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.TicketNumber);
            Actions.SetText(ticketNo, data.ItemArray[17].ToString());

            var joNo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.JobOrderNumber);
            Actions.SetText(joNo, data.ItemArray[18].ToString());

            var customerNo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.CustomerNumber);
            Actions.SetText(customerNo, data.ItemArray[19].ToString());

            var customerName = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[20].ToString());

            var customerFedId = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.CustomerId);
            Actions.SetText(customerFedId, data.ItemArray[21].ToString());

            var phone = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.Phone);
            Actions.SetText(phone, data.ItemArray[22].ToString());

            var address = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.Address);
            Actions.SetText(address, data.ItemArray[23].ToString());

            var city = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.City);
            Actions.SetText(city, data.ItemArray[24].ToString());

            var state = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.State);
            Actions.SetText(state, data.ItemArray[25].ToString());

            var zip = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.Zip);
            Actions.SetText(zip, data.ItemArray[26].ToString());

            var invoiceType = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceType);
            if (!string.IsNullOrEmpty(data.ItemArray[27].ToString()))
            DropDownActions.SelectDropdownByText(invoiceType, data.ItemArray[27].ToString());

            var status = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
            DropDownActions.SelectDropdownByText(status, data.ItemArray[28].ToString());

            var invoiceAmount = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceAmount);
            Actions.SetText(invoiceAmount, data.ItemArray[29].ToString());

            var balanceAmount = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.BalanceAmount);
            Actions.SetText(balanceAmount, data.ItemArray[30].ToString());

            var organization = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.Organization);
            if (!string.IsNullOrEmpty(data.ItemArray[31].ToString()))
            DropDownActions.SelectDropdownByText(organization, data.ItemArray[31].ToString());

            var invoiceFrom = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[32].ToString()))
            DropDownActions.SelectDropdownByText(invoiceFrom, data.ItemArray[32].ToString());

            var invoiceTo = Actions.GetWindowChild(searchWindow, InvoiceSearchConstants.InvoiceTo);
            if (!string.IsNullOrEmpty(data.ItemArray[33].ToString()))
            DropDownActions.SelectDropdownByText(invoiceTo, data.ItemArray[33].ToString());


        }

        public static bool SelectInvoiceRecord(string invoiceNo)
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Exists)
            {
                var label = Actions.GetWindowChild(searchWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var creditCard = TableActions.OpenRecordFromTable(searchWindow, "_grdSearchResult",
                    "Invoice Number", invoiceNo);
                return creditCard;
            }
            return false;
        }

        #endregion

        #region Invoice Transactions Search Methods

        public static void EnterTransactionSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var transactionfrom = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.InvoiceFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[35].ToString())) 
            DropDownActions.SelectDropdownByText(transactionfrom,data.ItemArray[35].ToString());

            var transactionto = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.InvoiceTo);
            if (!string.IsNullOrEmpty(data.ItemArray[36].ToString()))
                DropDownActions.SelectDropdownByText(transactionto, data.ItemArray[36].ToString());

            var labProInvoice = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.LabProInvoiceNo);
            Actions.SetText(labProInvoice,data.ItemArray[37].ToString());

            var billingNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.BillNo);
            Actions.SetText(billingNo, data.ItemArray[38].ToString());

            var customerNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CustomerNo);
            Actions.SetText(customerNo, data.ItemArray[39].ToString());

            var customerName = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[40].ToString());

            var address = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.Address);
            Actions.SetText(address, data.ItemArray[41].ToString());

            var fedId = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CustomerId);
            Actions.SetText(fedId, data.ItemArray[42].ToString());

            var phone = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.Phone);
            Actions.SetText(phone, data.ItemArray[43].ToString());

            var invoiceNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.EllisInvoiceNo);
            Actions.SetText(invoiceNo, data.ItemArray[44].ToString());

            var ticketNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.TicketNo);
            Actions.SetText(ticketNo, data.ItemArray[45].ToString());

            var jobOrderNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.JobOrderNo);
            Actions.SetText(jobOrderNo, data.ItemArray[46].ToString());

            var district = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.District);
            if (!string.IsNullOrEmpty(data.ItemArray[47].ToString()))
                DropDownActions.SelectDropdownByText(district, data.ItemArray[47].ToString());

            var branch = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[48].ToString()))
                DropDownActions.SelectDropdownByText(branch, data.ItemArray[48].ToString());
            
            var orderBy = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.OrderBy);
            if (!string.IsNullOrEmpty(data.ItemArray[51].ToString()))
                DropDownActions.SelectDropdownByText(orderBy, data.ItemArray[51].ToString());
          
            # region Credit

            var creditChkBox = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CreditsChkBox);
            if (!string.IsNullOrEmpty(data.ItemArray[49].ToString()))
            Actions.SetCheckBox((WinCheckBox)creditChkBox, data.ItemArray[49].ToString());

            if (data.ItemArray[49].ToString().Equals("TRUE"))
            {
                var creditStatus = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CreditStatus);
                if (!string.IsNullOrEmpty(data.ItemArray[52].ToString()))
                    DropDownActions.SelectDropdownByText(creditStatus, data.ItemArray[52].ToString());

                var creditType = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CreditType);
                if (!string.IsNullOrEmpty(data.ItemArray[53].ToString()))
                    DropDownActions.SelectDropdownByText(creditType, data.ItemArray[53].ToString());

                var amount = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.CreditAmt);
                Actions.SetText(amount, data.ItemArray[54].ToString());
            }
            
            # endregion

            # region Payment

            var paymentChkBox = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.PaymentsChkBox);
            if (!string.IsNullOrEmpty(data.ItemArray[50].ToString()))
            Actions.SetCheckBox((WinCheckBox)paymentChkBox, data.ItemArray[50].ToString());

            if (data.ItemArray[50].ToString().Equals("TRUE"))
            {
                var paymentRefNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.PaymentRefNo);
                Actions.SetText(paymentRefNo, data.ItemArray[55].ToString());

                var paymentType = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.PaymentType);
                if (!string.IsNullOrEmpty(data.ItemArray[56].ToString()))
                    DropDownActions.SelectDropdownByText(paymentType, data.ItemArray[56].ToString());

                var batchNo = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.BatchNo);
                Actions.SetText(batchNo, data.ItemArray[57].ToString());

                var paymentamt = Actions.GetWindowChild(searchWindow, InvoiceTransactionsConstants.PaymentAmt);
                Actions.SetText(paymentamt, data.ItemArray[58].ToString());
            }
           

            # endregion

          
        }

        public static bool SelectTransactionRecord(string transactionNumber)
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Exists)
            {
                var label = Actions.GetWindowChild(searchWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var transaction = TableActions.OpenRecordFromTable(searchWindow, "_grdSearchResult",
                    "LabPro Invoice Number", transactionNumber);
                return transaction;
            }
            return false;
        }

        #endregion

        #region Invoice Relationship Search Methods

        public static void EnterRelationshipSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var invoiceNo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceNo);
            Actions.SetText(invoiceNo,data.ItemArray[60].ToString());

            var billingNo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.BilingNo);
            Actions.SetText(billingNo, data.ItemArray[61].ToString());

            var ticketNo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.TicketNo);
            Actions.SetText(ticketNo, data.ItemArray[62].ToString());

            var jobOrderNo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.JobOrderNo);
            Actions.SetText(jobOrderNo, data.ItemArray[63].ToString());

            var customerName = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[64].ToString());

            var customerNo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.CustomerNo);
            Actions.SetText(customerNo, data.ItemArray[65].ToString());

            var customerId = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.CustomerId);
            Actions.SetText(customerId, data.ItemArray[66].ToString());

            var invoiceType = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceType);
            if (!string.IsNullOrEmpty(data.ItemArray[67].ToString()))
                DropDownActions.SelectDropdownByText(invoiceType, data.ItemArray[67].ToString());

            var invoiceStatus = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[68].ToString()))
                DropDownActions.SelectDropdownByText(invoiceStatus, data.ItemArray[68].ToString());

            var invoiceAmt = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceAmt);
            Actions.SetText(invoiceAmt, data.ItemArray[69].ToString());

            var balanceAmt = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.BalanceAmt);
            Actions.SetText(balanceAmt, data.ItemArray[70].ToString());

            var invoiceFrom = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[71].ToString()))
                DropDownActions.SelectDropdownByText(invoiceFrom, data.ItemArray[71].ToString());

            var invoiceTo = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.InvoiceTo);
            if (!string.IsNullOrEmpty(data.ItemArray[72].ToString()))
                DropDownActions.SelectDropdownByText(invoiceTo, data.ItemArray[72].ToString());

            var organization = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.Organization);
            if (!string.IsNullOrEmpty(data.ItemArray[73].ToString()))
                DropDownActions.SelectDropdownByText(organization, data.ItemArray[73].ToString());

            var organizationType = Actions.GetWindowChild(searchWindow, InvoiceRelationshipConstants.CollectionType);
            if (organizationType.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[74].ToString()))
                    DropDownActions.SelectDropdownByText(organizationType, data.ItemArray[74].ToString());
            }
            
        }

        #endregion

        #region Common Methods

        public static bool VerifySearchResultsWindowDisplayed()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                return true;
            }
            return false;
        }

        public static bool CloseSearchResultsWindow()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)searchResultsWindow);
                return true;
            }
            return false;
        }

        public static bool VerifySearchWindowDisplayed()
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool ClickCancelBtn()
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Exists)
            {
                MouseActions.ClickButton(searchWindow, "_cancelButton");
                return true;
            }
            return false;
        }

        public static bool ClickSearchBtn()
        {
            var searchWindow = GetSearchWindowProperties();
            if (searchWindow.Exists)
            {
                MouseActions.ClickButton(searchWindow, "_buttonSearch");
                return true;
            }
            return false;
        }

        public static bool VerifyDefaultDistrictSelected(string data)
        {
            var window = GetSearchWindowProperties();

            var cmb = Actions.GetWindowChild(window, "_cmbDistrict");
            var cmbBox = (WinComboBox) cmb;
            return cmbBox.SelectedItem.Equals(data);
        }

        public static bool VerifyInvoicingOrganizationIsNull()
        {
            var window = GetSearchWindowProperties();

            var cmb = Actions.GetWindowChild(window, "cmbOrganization");
            var cmbBox = (WinComboBox) cmb;
            return string.IsNullOrEmpty(cmbBox.SelectedItem);
        }

        public static bool ClickOnRefineSearchBtn()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                MouseActions.ClickButton(searchResultsWindow, "_btnRefineSearch");
                return true;
            }
            return false;
        }

        public static bool ClickOnPrintBtn()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                MouseActions.ClickButton(searchResultsWindow, "_btnPrint");
                SendKeys.SendWait("%{F4}");
                return true;
            }
            return false;
        }

        public static bool ClickOnExportBtn()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                MouseActions.ClickButton(searchResultsWindow, "_btnExport");
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        public class BillingSearchConstants
        {
            public const string CustomerName = "_txtCustomerName";
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string DispatchStart = "_dtpStartDate";
            public const string DispatchEnd = "_dtpEndDate";
            public const string District = "_ddlDistrict";
            public const string Branch = "_ddlBranch";
        }

        public class CreditCardSearchConstants
        {
            public const string CreditCard = "_txtCreditCardNo";
            public const string Transaction = "_cbTransactionSatus";
            public const string ProcessedFrom = "_calendarFrom";
            public const string ProcessedTo = "_calendarTo";
        }

        public class InvoiceSearchConstants
        {
            public const string InvoiceNumber = "_txtEllisInvoiceNumber";
            public const string BillingNumber = "_txtBillNumber";
            public const string TicketNumber = "_txtWorkTicketNumber";
            public const string JobOrderNumber = "_txtJobOrderNumber";
            public const string CustomerNumber = "_txtCustomerNumber";
            public const string CustomerName = "_txtCustomerName";
            public const string CustomerId = "_txtCustomerFEDID";
            public const string Phone = "_txtBillingPhone";
            public const string Address = "_txtBillingAddress";
            public const string City = "_txtBillingAddressCity";
            public const string State = "_txtBillingAddressState";
            public const string Zip = "_txtBillingAddressZip";
            public const string InvoiceType = "_cboType";
            public const string InvoiceStatus = "_cboStatus";
            public const string InvoiceAmount = "_txtInvoiceAmount";
            public const string BalanceAmount = "_txtBalanceDueAmount";
            public const string Organization = "cmbOrganization";
            public const string InvoiceFrom = "_calendarFrom";
            public const string InvoiceTo = "_calendarTo";
        }

        public class InvoiceTransactionsConstants
        {
            public const string InvoiceFrom = "_ccInvoiceDateFrom";
            public const string InvoiceTo = "_ccInvoiceDateTo";
            public const string LabProInvoiceNo = "_txtLabProInvoiceNumber";
            public const string BillNo = "_txtBillNumber";
            public const string CustomerNo = "_txtCustomerNumber";
            public const string CustomerName = "_txtCustomerName";
            public const string Address = "_txtCustBillingAddr";
            public const string CustomerId = "_txtCustFedID";
            public const string Phone = "_txtCustomerPhone";
            public const string EllisInvoiceNo = "_txtEllisInvoiceNumber";
            public const string TicketNo = "_txtWorkTicketNumber";
            public const string JobOrderNo = "_txtJobOrderNumber";
            public const string District = "_cmbDistrict";
            public const string Branch = "_cmbBranch";
            public const string CreditsChkBox = "_cboxCredits";
            public const string PaymentsChkBox = "_cboxPayments";
            public const string OrderBy = "_cboOrderBy";
            public const string CreditStatus = "_cboCreditRequestStatus";
            public const string CreditType = "_cboCreditRequestType";
            public const string CreditAmt = "_txtCreditAmount";
            public const string PaymentRefNo = "_txtPaymtRefNo";
            public const string PaymentType = "_cboPaymentType";
            public const string BatchNo = "_txtBatchNumber";
            public const string PaymentAmt = "_txtPaymentAmount";
        }

        public class InvoiceRelationshipConstants
        {
            public const string InvoiceNo = "_txtInvoiceNumber";
            public const string BilingNo = "ultraTextEditor1";
            public const string TicketNo = "_txtWorkTicketNumber";
            public const string JobOrderNo = "_txtJobOrder";
            public const string CustomerName = "txtCustomerName";
            public const string CustomerNo = "_txtCustomerNumber";
            public const string CustomerId = "_txtCustomerFEDID";
            public const string InvoiceType = "_cboType";
            public const string InvoiceStatus = "_cboStatus";
            public const string InvoiceAmt = "_txtInvoiceAmount";
            public const string BalanceAmt = "_txtBalanceDueAmount";
            public const string InvoiceFrom = "_calendarFrom";
            public const string InvoiceTo = "_calendarTo";
            public const string Organization = "cmbOrganization";
            public const string CollectionType = "cmbCollectionType";
            
        }

        #endregion
    }
}