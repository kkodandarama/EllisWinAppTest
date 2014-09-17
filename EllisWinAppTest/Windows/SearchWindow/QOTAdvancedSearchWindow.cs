using System.Data;
using System.Linq;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.SearchWindow
{
    public class QOTAdvancedSearchWindow : AppContext
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

        private static UITestControl GetPaymentDetailsWindowProperties()
        {
            var paymentDetailsWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Payment Details" });
            return paymentDetailsWindow;
        }

        private static UITestControl GetDispatchWindowProperties()
        {
            var dispatchWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return dispatchWindow;
        }

        private static UITestControl GetQuoteProfileWindowProperties()
        {
            var quoteProfileWindow =
                App.Container.SearchFor<WinWindow>(new { Name = "Quote Profile" });
            return quoteProfileWindow;
        }

        #endregion

        #region Check Register Search Methods

        public static void EnterCheckRegisterSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var paymentType = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.PaymentType);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                DropDownActions.SelectDropdownByText(paymentType, data.ItemArray[3].ToString());

            var paymentStatus = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.PaymentStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                DropDownActions.SelectDropdownByText(paymentStatus, data.ItemArray[4].ToString());

            var paymentFrom = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.PaymentFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                DropDownActions.SelectDropdownByText(paymentFrom, data.ItemArray[5].ToString());

            var paymentTo = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.PaymentTo);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                DropDownActions.SelectDropdownByText(paymentTo, data.ItemArray[6].ToString());

            var district = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.District);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                DropDownActions.SelectDropdownByText(district, data.ItemArray[7].ToString());

            var branch = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                DropDownActions.SelectDropdownByText(branch, data.ItemArray[8].ToString());

            var deductionType = Actions.GetWindowChild(searchWindow, CheckRegisterConstants.DeductionType);
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                DropDownActions.SelectDropdownByText(deductionType, data.ItemArray[9].ToString());
        }

        public static bool SelectCheckRegisterRecord(string documentNo)
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
                var document = TableActions.OpenRecordFromTable(searchWindow, "_grdSearchResult",
                    "Document Number", documentNo);
                return document;
            }
            return false;
        }

        #endregion

        #region Work Ticket Search Methods

        public static void EnterWorkTicketSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var ticketNo = Actions.GetWindowChild(searchWindow, WorkTicketConstants.TicketNo);
            Actions.SetText(ticketNo,data.ItemArray[11].ToString());

            var jobOrderNo = Actions.GetWindowChild(searchWindow, WorkTicketConstants.JobOrderNo);
            Actions.SetText(jobOrderNo, data.ItemArray[12].ToString());

            var state = Actions.GetWindowChild(searchWindow, WorkTicketConstants.State);
            if(!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                DropDownActions.SelectDropdownByText(state,data.ItemArray[13].ToString());

            var city     = Actions.GetWindowChild(searchWindow, WorkTicketConstants.City);
            Actions.SetText(city, data.ItemArray[14].ToString());

            var customerName = Actions.GetWindowChild(searchWindow, WorkTicketConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[15].ToString());

            var workerName = Actions.GetWindowChild(searchWindow, WorkTicketConstants.WorkerName);
            Actions.SetText(workerName, data.ItemArray[16].ToString());

            var ssn = Actions.GetWindowChild(searchWindow, WorkTicketConstants.Ssn);
            Actions.SetText(ssn, data.ItemArray[17].ToString());

            var branchNo = Actions.GetWindowChild(searchWindow, WorkTicketConstants.BranchNo);
            Actions.SetText(branchNo, data.ItemArray[18].ToString());

            var date = Actions.GetWindowChild(searchWindow, WorkTicketConstants.DispatchDate);
            if (!string.IsNullOrEmpty(data.ItemArray[19].ToString()))
                DropDownActions.SelectDropdownByText(date, data.ItemArray[19].ToString());

            var startTime = Actions.GetWindowChild(searchWindow, WorkTicketConstants.StartTime);
            SendKeys.SendWait("{HOME}");
            Actions.SetText(startTime, data.ItemArray[20].ToString());
        }

        public static bool SelectWorkTicketRecord(string jobOrderNo)
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
                var ticket = TableActions.OpenRecordFromTable(searchWindow, "_grdSearchResult",
                    "Work Ticket Number", jobOrderNo);
                return ticket;
            }
            return false;
        }

        #endregion

        #region Quote Search Methods

        public static void EnterQuoteSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var requester = Actions.GetWindowChild(searchWindow, QuoteConstants.RequesterName);
            Actions.SetText(requester,data.ItemArray[22].ToString());

            var quoteNo = Actions.GetWindowChild(searchWindow, QuoteConstants.QuoteNo);
            Actions.SetText(quoteNo, data.ItemArray[23].ToString());

            var createDt = Actions.GetWindowChild(searchWindow, QuoteConstants.CreateDt);
            if (!string.IsNullOrEmpty(data.ItemArray[24].ToString()))
                DropDownActions.SelectDropdownByText(createDt,data.ItemArray[24].ToString());

            var effectiveDt = Actions.GetWindowChild(searchWindow, QuoteConstants.EffectiveDt);
            if (!string.IsNullOrEmpty(data.ItemArray[25].ToString()))
                DropDownActions.SelectDropdownByText(effectiveDt, data.ItemArray[25].ToString());

            var expirationDt = Actions.GetWindowChild(searchWindow, QuoteConstants.ExpirationDt);
            if (!string.IsNullOrEmpty(data.ItemArray[26].ToString()))
                DropDownActions.SelectDropdownByText(expirationDt, data.ItemArray[26].ToString());

            var projectName = Actions.GetWindowChild(searchWindow, QuoteConstants.ProjectNo);
            Actions.SetText(projectName, data.ItemArray[27].ToString());

            var organizationType = Actions.GetWindowChild(searchWindow, QuoteConstants.OraganizationType);
            if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                DropDownActions.SelectDropdownByText(organizationType, data.ItemArray[28].ToString());

            var organization = Actions.GetWindowChild(searchWindow, QuoteConstants.QuoteOrganization);
            if (!string.IsNullOrEmpty(data.ItemArray[29].ToString()))
                DropDownActions.SelectDropdownByText(organization, data.ItemArray[29].ToString());

            var user = Actions.GetWindowChild(searchWindow, QuoteConstants.Username);
            Actions.SetText(user, data.ItemArray[30].ToString());

            var prevailingChk = Actions.GetWindowChild(searchWindow, QuoteConstants.PrevailingChkBox);
            if (!string.IsNullOrEmpty(data.ItemArray[31].ToString()))
                Actions.SetCheckBox((WinCheckBox)prevailingChk, data.ItemArray[31].ToString());

            #region Wrap Check Box

            var wrapChk = Actions.GetWindowChild(searchWindow, QuoteConstants.WrapChkBox);
            if (!string.IsNullOrEmpty(data.ItemArray[32].ToString()))
                Actions.SetCheckBox((WinCheckBox)wrapChk, data.ItemArray[32].ToString());

            if (data.ItemArray[32].ToString().Equals("TRUE"))
            {
                var evaluationDt = Actions.GetWindowChild(searchWindow, QuoteConstants.WrapEvaluationDt);
                if (!string.IsNullOrEmpty(data.ItemArray[33].ToString()))
                    DropDownActions.SelectDropdownByText(evaluationDt, data.ItemArray[33].ToString());

                var status = Actions.GetWindowChild(searchWindow, QuoteConstants.WrapStatus);
                if (!string.IsNullOrEmpty(data.ItemArray[34].ToString()))
                    DropDownActions.SelectDropdownByText(status, data.ItemArray[34].ToString());
            }

            #endregion

           
            var quoteType = Actions.GetWindowChild(searchWindow, QuoteConstants.QuoteType);
            if (!string.IsNullOrEmpty(data.ItemArray[35].ToString()))
                DropDownActions.SelectDropdownByText(quoteType, data.ItemArray[35].ToString());

            var workCat = Actions.GetWindowChild(searchWindow, QuoteConstants.WorkCategory);
            if (!string.IsNullOrEmpty(data.ItemArray[36].ToString()))
                DropDownActions.SelectDropdownByText(workCat, data.ItemArray[36].ToString());

            //var duties = Actions.GetWindowChild(searchWindow, QuoteConstants.JobDuties);
            //if (duties.Enabled)
            //{
            //    Actions.SetText(duties, data.ItemArray[37].ToString());
            //}

            var quoteStatus = Actions.GetWindowChild(searchWindow, QuoteConstants.QuoteStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[38].ToString()))
                DropDownActions.SelectDropdownByText(quoteStatus, data.ItemArray[37].ToString());

            var compChk = Actions.GetWindowChild(searchWindow, QuoteConstants.CompCodeChkBox);
            if (!string.IsNullOrEmpty(data.ItemArray[39].ToString()))
            Actions.SetCheckBox((WinCheckBox)compChk, data.ItemArray[39].ToString());

            
        }

        public static bool SelectQuoteRecord(string quoteNo)
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                var label = Actions.GetWindowChild(searchResultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var quote = TableActions.OpenRecordFromTable(searchResultsWindow, "_grdSearchResult", "Quote Number",
                    quoteNo);
                return quote;
            }
            return false;
        }

        #endregion

        #region Dispatch Search Methods

        public static void EnterDispatchSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var joborderNo = Actions.GetWindowChild(searchWindow, DispatchConstants.JobOrderNo);
            Actions.SetText(joborderNo,data.ItemArray[41].ToString());

            var orderStatus = Actions.GetWindowChild(searchWindow, DispatchConstants.OrderStatus);
            if(!string.IsNullOrEmpty(data.ItemArray[42].ToString()))
            DropDownActions.SelectDropdownByText(orderStatus,data.ItemArray[42].ToString());

            var jobDuty = Actions.GetWindowChild(searchWindow, DispatchConstants.JobDuty);
            if (!string.IsNullOrEmpty(data.ItemArray[43].ToString()))
            DropDownActions.SelectDropdownByText(jobDuty, data.ItemArray[43].ToString());

            var orderFrom = Actions.GetWindowChild(searchWindow, DispatchConstants.OrderFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[44].ToString()))
            DropDownActions.SelectDropdownByText(orderFrom, data.ItemArray[44].ToString());

            var orderTo = Actions.GetWindowChild(searchWindow, DispatchConstants.OrderTo);
            if (!string.IsNullOrEmpty(data.ItemArray[45].ToString()))
            DropDownActions.SelectDropdownByText(orderTo, data.ItemArray[45].ToString());

            var startTime = Actions.GetWindowChild(searchWindow, DispatchConstants.StartTime);
            Actions.SetText(startTime, data.ItemArray[46].ToString());

            var workerFrom = Actions.GetWindowChild(searchWindow, DispatchConstants.WorkerFrom);
            Actions.SetText(workerFrom, data.ItemArray[47].ToString());

            var workerTo = Actions.GetWindowChild(searchWindow, DispatchConstants.WorkerTo);
            Actions.SetText(workerTo, data.ItemArray[48].ToString());

            var state = Actions.GetWindowChild(searchWindow, DispatchConstants.State);
            if (!string.IsNullOrEmpty(data.ItemArray[49].ToString()))
                DropDownActions.SelectDropdownByText(state, data.ItemArray[49].ToString());

            var city = Actions.GetWindowChild(searchWindow, DispatchConstants.City);
            if (!string.IsNullOrEmpty(data.ItemArray[50].ToString()))
                DropDownActions.SelectDropdownByText(city, data.ItemArray[50].ToString());

            var branch = Actions.GetWindowChild(searchWindow, DispatchConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[51].ToString()))
                DropDownActions.SelectDropdownByText(branch, data.ItemArray[51].ToString());

            var organization = Actions.GetWindowChild(searchWindow, DispatchConstants.Organization);
            if (!string.IsNullOrEmpty(data.ItemArray[52].ToString()))
                DropDownActions.SelectDropdownByText(organization, data.ItemArray[52].ToString());
        }

        public static bool SelectDispatchRecord(string customerName)
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                var label = Actions.GetWindowChild(searchResultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var dispatch = TableActions.OpenRecordFromTable(searchResultsWindow, "_grdSearchResult", "Customer",
                    customerName);
                return dispatch;
            }
            return false;
        }

        #endregion

        #region Job Order Search Methods

        public static void EnterJobOrderSearchData(DataRow data)
        {
            var searchWindow = GetSearchWindowProperties();

            var joborderNo = Actions.GetWindowChild(searchWindow, JobOrderConstants.JobOrderNo);
            Actions.SetText(joborderNo,data.ItemArray[54].ToString());

            var status = Actions.GetWindowChild(searchWindow, JobOrderConstants.JobOrderStatus);
            if(!string.IsNullOrEmpty(data.ItemArray[55].ToString()))
                DropDownActions.SelectDropdownByText(status,data.ItemArray[55].ToString());

            var quoteNo = Actions.GetWindowChild(searchWindow, JobOrderConstants.QuoteNo);
            Actions.SetText(quoteNo, data.ItemArray[56].ToString());

            var workerCode = Actions.GetWindowChild(searchWindow, JobOrderConstants.CompensationCode);
            Actions.SetText(workerCode, data.ItemArray[57].ToString());

            var customerName = Actions.GetWindowChild(searchWindow, JobOrderConstants.CustomerName);
            Actions.SetText(customerName, data.ItemArray[58].ToString());

            var customerId = Actions.GetWindowChild(searchWindow, JobOrderConstants.CustomerId);
            Actions.SetText(customerId, data.ItemArray[59].ToString());

            var placedBy = Actions.GetWindowChild(searchWindow, JobOrderConstants.Customer);
            Actions.SetText(placedBy, data.ItemArray[60].ToString());

            var poNo = Actions.GetWindowChild(searchWindow, JobOrderConstants.PoNo);
            Actions.SetText(poNo, data.ItemArray[61].ToString());

            var costCenter = Actions.GetWindowChild(searchWindow, JobOrderConstants.CostCenter);
            Actions.SetText(costCenter, data.ItemArray[62].ToString());

            var projectName = Actions.GetWindowChild(searchWindow, JobOrderConstants.ProjectName);
            Actions.SetText(projectName, data.ItemArray[63].ToString());

            var jobDuty = Actions.GetWindowChild(searchWindow, JobOrderConstants.JobDuty);
            if (!string.IsNullOrEmpty(data.ItemArray[64].ToString()))
                DropDownActions.SelectDropdownByText(jobDuty, data.ItemArray[64].ToString());

            var state = Actions.GetWindowChild(searchWindow, JobOrderConstants.State);
            if (!string.IsNullOrEmpty(data.ItemArray[65].ToString()))
                DropDownActions.SelectDropdownByText(state, data.ItemArray[65].ToString());

            var city = Actions.GetWindowChild(searchWindow, JobOrderConstants.City);
            if (!string.IsNullOrEmpty(data.ItemArray[66].ToString()))
                DropDownActions.SelectDropdownByText(city, data.ItemArray[66].ToString());

            var codRadio = Actions.GetWindowChild(searchWindow, JobOrderConstants.RadioCod);
            if (!string.IsNullOrEmpty(data.ItemArray[67].ToString()))
            {
                var radio = codRadio.Container.SearchFor<WinRadioButton>(new { Name = data.ItemArray[67].ToString() });
                Actions.SelectRadioButton(radio);
            }

            var prevailRadio = Actions.GetWindowChild(searchWindow, JobOrderConstants.RadioPrevailing);
            if (!string.IsNullOrEmpty(data.ItemArray[68].ToString()))
            {
                var radio = prevailRadio.Container.SearchFor<WinRadioButton>(new { Name = data.ItemArray[68].ToString() });
                Actions.SelectRadioButton(radio);
            }

            var wrapRadio = Actions.GetWindowChild(searchWindow, JobOrderConstants.RadioInsurance);
            if (!string.IsNullOrEmpty(data.ItemArray[69].ToString()))
            {
                var radio = wrapRadio.Container.SearchFor<WinRadioButton>(new { Name = data.ItemArray[69].ToString() });
                Actions.SelectRadioButton(radio);
            }

            var effectiveFrom = Actions.GetWindowChild(searchWindow, JobOrderConstants.EffectiveFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[70].ToString()))
                DropDownActions.SelectDropdownByText(effectiveFrom, data.ItemArray[70].ToString());

            var effectiveTo = Actions.GetWindowChild(searchWindow, JobOrderConstants.EffectiveTo);
            if (!string.IsNullOrEmpty(data.ItemArray[71].ToString()))
                DropDownActions.SelectDropdownByText(effectiveTo, data.ItemArray[71].ToString());

            var expirationFrom = Actions.GetWindowChild(searchWindow, JobOrderConstants.ExpirationFrom);
            if (!string.IsNullOrEmpty(data.ItemArray[72].ToString()))
                DropDownActions.SelectDropdownByText(expirationFrom, data.ItemArray[72].ToString());

            var expirationTo = Actions.GetWindowChild(searchWindow, JobOrderConstants.ExpirationTo);
            if (!string.IsNullOrEmpty(data.ItemArray[73].ToString()))
                DropDownActions.SelectDropdownByText(expirationTo, data.ItemArray[73].ToString());

            var organization = Actions.GetWindowChild(searchWindow, JobOrderConstants.Branch);
            if (!string.IsNullOrEmpty(data.ItemArray[74].ToString()))
                DropDownActions.SelectDropdownByText(organization, data.ItemArray[74].ToString());

            var user = Actions.GetWindowChild(searchWindow, JobOrderConstants.UserName);
            Actions.SetText(user, data.ItemArray[75].ToString());

        }

        public static bool SelectJobOrderRecord(string jobOrderNo)
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Exists)
            {
                var label = Actions.GetWindowChild(searchResultsWindow, "CasePartTitleLabel");
                var text = label.GetProperty("Name").ToString().Contains("No records found");
                if (text)
                {
                    return false;
                }
                var joborder = TableActions.OpenRecordFromTable(searchResultsWindow, "_grdSearchResult", "Job Order#",
                    jobOrderNo);
                if(joborder)
                    MouseActions.ClickButton(JobOrderWindow.JobOrderWindow.GetJobOrderProfileProperties(), "btnClose");
                return joborder;
            }
            return false;
        }

        #endregion

        #region Common Methods

        public static bool VerifyQuoteProfileWindowDisplayed()
        {
            var quoteProfileWindow = GetQuoteProfileWindowProperties();
            if (quoteProfileWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool CloseQuoteProfileWindow()
        {
            var quoteProfileWindow = GetQuoteProfileWindowProperties();
            if (quoteProfileWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)quoteProfileWindow);
                return true;
            }
            return false;
        }

        public static bool VerifyDispatchWindowDisplayed()
        {
            var dispatchWindow = GetDispatchWindowProperties();
            if (dispatchWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool CloseDispatchWindow()
        {
            var dispatchWindow = GetDispatchWindowProperties();
            if (dispatchWindow.Exists)
            {
                TitlebarActions.ClickClose((WinWindow)dispatchWindow);
                return true;
            }
            return false;
        }

        public static bool VerifySearchResultsWindowDisplayed()
        {
            var searchResultsWindow = GetSearchResultsWindowProperties();
            if (searchResultsWindow.Enabled)
            {
                return true;
            }
            return false;
        }

        public static bool VerifyPaymentDetailsWindowDisplayed()
        {
            var paymentDetailsWindow = GetPaymentDetailsWindowProperties();
            if (paymentDetailsWindow.Enabled)
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

        public static bool ClickOkBtnPayment()
        {
            var paymentDetailsWindow = GetPaymentDetailsWindowProperties();
            if (paymentDetailsWindow.Exists)
            {
                MouseActions.ClickButton(paymentDetailsWindow, "btnOK");
                return true;
            }
            return false;
        }

        public static bool ClickVoidBtnPayment()
        {
            var paymentDetailsWindow = GetPaymentDetailsWindowProperties();
            if (paymentDetailsWindow.Exists)
            {
                MouseActions.ClickButton(paymentDetailsWindow, "btnVoid");
                return true;
            }
            return false;
        }

        #endregion

        #region Controls

        public class CheckRegisterConstants
        {
            public const string PaymentType = "cmbPaymentType";
            public const string PaymentStatus = "cmbPaymentStatus";
            public const string PaymentFrom = "cmbDateRangeFrom";
            public const string PaymentTo = "cmbDateRangeTo";
            public const string District = "cmbDistrict";
            public const string Branch = "cmbBranch";
            public const string DeductionType = "cmbDeductionType";
        }

        public class WorkTicketConstants
        {
            public const string TicketNo = "mskWorkTicketNumber";
            public const string JobOrderNo = "txtJobOrderNumber";
            public const string State = "cmbState";
            public const string City = "txtCity";
            public const string CustomerName = "txtCustomerName";
            public const string WorkerName = "txtWorkerName";
            public const string Ssn = "mskSSN";
            public const string BranchNo = "txtBranchNumber";
            public const string DispatchDate = "_udtDispatchDate";
            public const string StartTime = "mskDispatchStartTime";
        }

        public class QuoteConstants
        {
            public const string RequesterName = "_quoteRequesterName";
            public const string QuoteNo = "_quoteNumber";
            public const string CreateDt = "_quoteCreatedDate";
            public const string EffectiveDt = "_quoteEffectiveDate";
            public const string ExpirationDt = "_quoteExpirationDate";
            public const string ProjectNo = "_customerProjectNumber";
            public const string OraganizationType = "_organizationType";
            public const string QuoteOrganization = "_quotedByOrganization";
            public const string Username = "_quotedByUserName";
            public const string PrevailingChkBox = "_prevailingWageQuote";
            public const string WrapChkBox = "_wrapUpQuote";
            public const string WrapEvaluationDt = "_wrapUpEvaluationDate";
            public const string WrapStatus = "_wrapStatus";
            public const string QuoteType = "_quoteType";
            public const string WorkCategory = "_workCategories";
            public const string JobDuties = "_jobDuties";
            public const string QuoteStatus = "_quoteStatus";
            public const string CompCodeChkBox = "_otherThanPrimaryCompCode";
        }

        public class DispatchConstants
        {
            public const string JobOrderNo = "mskJobOrderNumber";
            public const string OrderStatus = "cmbOrderDetailsStatus";
            public const string JobDuty = "cmbJobDuty";
            public const string OrderFrom = "dtpOrderDetailsFromDate";
            public const string OrderTo = "dtpOrderDetailsToDate";
            public const string StartTime = "mskStartTime";
            public const string WorkerFrom = "mskTotalWorkersFrom";
            public const string WorkerTo = "mskTotalWorkersTo";
            public const string State = "cmbState";
            public const string City = "cmbCity";
            public const string Branch = "cmbAssignedTo";
            public const string Organization = "cmbOwnerBranch";
            
        }

        public class JobOrderConstants
        {
            public const string JobOrderNo = "mskJobOrderNumber";
            public const string JobOrderStatus = "cmbJobOrderStatus";
            public const string QuoteNo = "mskQuoteNumber";
            public const string CompensationCode = "txtCompensationCode";
            public const string CustomerName = "txtCustomerName";
            public const string CustomerId = "mskFedID";
            public const string Customer = "txtPlacedByCustomer";
            public const string PoNo = "txtPONumber";
            public const string CostCenter = "txtCustomerCostCenter";
            public const string ProjectName = "txtCustomerProjectName";
            public const string JobDuty = "cmbJobDuty";
            public const string State = "cmbState";
            public const string City = "cmbCity";
            public const string RadioCod = "optCOD";
            public const string RadioPrevailing = "optPrevailingWageJob";
            public const string RadioInsurance = "optWrapInsurance";
            public const string EffectiveFrom = "dtpFromDate";
            public const string EffectiveTo = "dtpToDate";
            public const string ExpirationFrom = "dtpExpirationDateFrom";
            public const string ExpirationTo = "dtpExpirationDateTo";
            public const string Branch = "cmbOwnerBranch";
            public const string UserName = "txtEllisUserName";

        }

        #endregion
    }
}
