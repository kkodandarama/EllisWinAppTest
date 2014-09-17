using System;
using System.Data;
using System.Linq;
using System.Runtime.Remoting;
using System.Windows.Automation;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.InteropServices;

namespace EllisWinAppTest.Windows.PayrollWindow
{
    public class CreateNewOrderWindow : AppContext
    {
        #region Window Properties

        private static UITestControl GetCreateNewOrderWindowProperties()
        {

            return Actions.GetWindowProperties(App, "Create New Order");

        }

        private static UITestControl GetAttachNewOrderImageWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Attach new Order Image");
        }

        private static UITestControl GetSearchOrderWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Search Order");
        }

        private static UITestControl GetOrderDetailsWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Order Details");
        }

        private static UITestControl GetAddNoteWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Add Note");
        }

        private static UITestControl GetSearchAgencyWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Search Agency");
        }

        private static UITestControl GetUpdateAgencyWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Update Agency");
        }

        private static UITestControl GetAnswerLetterWindowProperties()
        {
            return Actions.GetWindowProperties(App, "View and Print Answer Letters");
        }

        private static UITestControl GetRateSheetWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Rate Sheet");
        }

        private static UITestControl GetRateSheetDateWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Rate Sheet Date");
        }

        private static UITestControl GetprintWindowProperties()
        {
            var answerWindow = GetAnswerLetterWindowProperties();
            var printWindow = answerWindow.Container.SearchFor<WinWindow>(new { Name = "Print" });
            return printWindow;
        }

        private static AutomationElement GetPaymentSearchCloseWindowButton()
        {
            var paymentSearch = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Payment Search"));
            var closedButton = paymentSearch.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Close"));
            
            return closedButton;
        }

        private static UITestControl GetOverTimeWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Overtime");
        }

        private static UITestControl GetSaveSuccessWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Save Success");
        }

        private static UITestControl GetNoResultsFoundWindowProperties()
        {
            return Actions.GetWindowProperties(App, "No Result found");
        }

        private static UITestControl GetSearchPaymentsWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Payment Search");
        }

        private static UITestControl GetClearedCheckGatewayWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Cleared Check Gateway");
        }

        private static UITestControl GetSchedulePaymentsWindowProperties()
        {
            return Actions.GetWindowProperties(App, "Schedule Payments");
        }

        #endregion

        #region Create New Order Methods

        public static void ClickAttachImageButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.AttachImage);
        }

        public static bool VerifyAttachNewOrderImageWindowDisplayed()
        {
            var window = GetAttachNewOrderImageWindowProperties();
            return window.Exists;
        }

        public static bool VerifyCreatNeworderWindowDisplayed()
        {
            var createNewOrder = GetCreateNewOrderWindowProperties();
            return createNewOrder.Exists;
        }

        public static void ClickNewOrderImageCancelButton()
        {
            var window = GetAttachNewOrderImageWindowProperties();
            var button = window.Container.SearchFor<WinButton>(new { ControlId = "2" });
            button.SetFocus();
            //Actions.SendEnter();
            MouseActions.Click(button);
            //TitlebarActions.ClickClose((WinWindow)window);
        }

        public static void ClickAcceptButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.AcceptBtn);
        }

        public static void ClickTreatAsNewButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.TreatAsNewBtn);
        }

        public static void ClickResetButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.ResetBtn);
        }

        public static void ClickSearchOrderButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.SearchOrderBtn);
        }

        public static void ClickSearchAgencyButton()
        {
            MouseActions.ClickButton(GetCreateNewOrderWindowProperties(), NewOrderControls.SearchAgencyBtn);
        }

        public static void EnterDataInCreateOrder(DataRow data)
        {
            var createOrderWindow = GetCreateNewOrderWindowProperties();

            var recipient = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Recipient);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                DropDownActions.SelectDropdownByText(recipient, data.ItemArray[3].ToString());

            var ssn = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Ssn);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                Actions.SetText(ssn, data.ItemArray[4].ToString());

            var firstname = Actions.GetWindowChild(createOrderWindow, NewOrderControls.FirstName);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                Actions.SetText(firstname, data.ItemArray[5].ToString());

            var middlename = Actions.GetWindowChild(createOrderWindow, NewOrderControls.MiddleName);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                Actions.SetText(middlename, data.ItemArray[6].ToString());

            var lastname = Actions.GetWindowChild(createOrderWindow, NewOrderControls.LastName);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                Actions.SetText(lastname, data.ItemArray[7].ToString());

            var dob = Actions.GetWindowChild(createOrderWindow, NewOrderControls.BirthDate);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                DropDownActions.SelectDropdownByText(dob, data.ItemArray[8].ToString());

            var firstworkDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.FirstWorkDt);
            if (firstworkDt.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                    Actions.SetText(firstworkDt, data.ItemArray[9].ToString());
            }

            var status = Actions.GetWindowChild(createOrderWindow, NewOrderControls.WorkerStatus);
            if (status.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                    Actions.SetText(status, data.ItemArray[10].ToString());
            }

            var appliedWorkDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.AppliedForWork);
            if (appliedWorkDt.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
                    Actions.SetText(appliedWorkDt, data.ItemArray[11].ToString());
            }

            var martialStatus = Actions.GetWindowChild(createOrderWindow, NewOrderControls.MartialStatus);
            if (martialStatus.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                    Actions.SetText(martialStatus, data.ItemArray[12].ToString());
            }

            var lastworkDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.LastWorkDt);
            if (lastworkDt.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
                    Actions.SetText(lastworkDt, data.ItemArray[13].ToString());
            }

            var agencyNo = Actions.GetWindowChild(createOrderWindow, NewOrderControls.AgencyNo);
            if (!string.IsNullOrEmpty(data.ItemArray[14].ToString()))
                Actions.SetText(agencyNo, data.ItemArray[14].ToString());

            var apVendor = Actions.GetWindowChild(createOrderWindow, NewOrderControls.ApVendor);
            if (!string.IsNullOrEmpty(data.ItemArray[15].ToString()))
                Actions.SetText(apVendor, data.ItemArray[15].ToString());

            var agencyName = Actions.GetWindowChild(createOrderWindow, NewOrderControls.AgencyName);
            if (!string.IsNullOrEmpty(data.ItemArray[16].ToString()))
                Actions.SetText(agencyName, data.ItemArray[16].ToString());

            var agencyZip = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Zip);
            if (!string.IsNullOrEmpty(data.ItemArray[17].ToString()))
                Actions.SetText(agencyZip, data.ItemArray[17].ToString());

            var agencyCountry = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Country);
            if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
                DropDownActions.SelectDropdownByText(agencyCountry, data.ItemArray[18].ToString());

            var agencyState = Actions.GetWindowChild(createOrderWindow, NewOrderControls.State);
            if (!string.IsNullOrEmpty(data.ItemArray[19].ToString()))
                DropDownActions.SelectDropdownByText(agencyState, data.ItemArray[19].ToString());

            var orderType = Actions.GetWindowChild(createOrderWindow, NewOrderControls.TypeOfOrder);
            if (!string.IsNullOrEmpty(data.ItemArray[20].ToString()))
                DropDownActions.SelectDropdownByText(orderType, data.ItemArray[20].ToString());

            var primaryCase = Actions.GetWindowChild(createOrderWindow, NewOrderControls.PrimaryCaseNo);
            if (!string.IsNullOrEmpty(data.ItemArray[21].ToString()))
                Actions.SetText(primaryCase, data.ItemArray[21].ToString());

            var secondaryCase = Actions.GetWindowChild(createOrderWindow, NewOrderControls.SecondaryCaseNo);
            if (!string.IsNullOrEmpty(data.ItemArray[22].ToString()))
                Actions.SetText(secondaryCase, data.ItemArray[22].ToString());

            var originalorderDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.OriginalOrderDt);
            if (!string.IsNullOrEmpty(data.ItemArray[23].ToString()))
                DropDownActions.SelectDropdownByText(originalorderDt, data.ItemArray[23].ToString());

            var orderStatus = Actions.GetWindowChild(createOrderWindow, NewOrderControls.OrderStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[24].ToString()))
                DropDownActions.SelectDropdownByText(orderStatus, data.ItemArray[24].ToString());

            var releaseVerify = Actions.GetWindowChild(createOrderWindow, NewOrderControls.ReleaseVerify);
            if (releaseVerify.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[25].ToString()))
                    DropDownActions.SelectDropdownByText(releaseVerify, data.ItemArray[25].ToString());
            }

            var statusOrderDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.StatusOrderDt);
            if (statusOrderDt.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[26].ToString()))
                    DropDownActions.SelectDropdownByText(statusOrderDt, data.ItemArray[26].ToString());
            }

            var arrears = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Arrears);
            if (!string.IsNullOrEmpty(data.ItemArray[27].ToString()))
                DropDownActions.SelectDropdownByText(arrears, data.ItemArray[27].ToString());

            var expirationMethod = Actions.GetWindowChild(createOrderWindow, NewOrderControls.ExpirationMethodDt);
            if (!string.IsNullOrEmpty(data.ItemArray[28].ToString()))
                DropDownActions.SelectDropdownByText(expirationMethod, data.ItemArray[28].ToString());

            var expirationDt = Actions.GetWindowChild(createOrderWindow, NewOrderControls.ExpiredDt);
            if (!string.IsNullOrEmpty(data.ItemArray[29].ToString()))
                DropDownActions.SelectDropdownByText(expirationDt, data.ItemArray[29].ToString());

            var orderPriority = Actions.GetWindowChild(createOrderWindow, NewOrderControls.OrderPriority);
            if (orderPriority.Enabled)
            {
                if (!string.IsNullOrEmpty(data.ItemArray[30].ToString()))
                    Actions.SetText(orderPriority, data.ItemArray[30].ToString());
            }

            var deduct = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Deduct);
            if (!string.IsNullOrEmpty(data.ItemArray[31].ToString()))
                Actions.SetText(deduct, data.ItemArray[31].ToString());

            var per = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Per);
            if (!string.IsNullOrEmpty(data.ItemArray[32].ToString()))
                DropDownActions.SelectDropdownByText(per, data.ItemArray[32].ToString());

            var totalMax = Actions.GetWindowChild(createOrderWindow, NewOrderControls.TotalMax);
            if (!string.IsNullOrEmpty(data.ItemArray[33].ToString()))
                Actions.SetText(totalMax, data.ItemArray[33].ToString());

            var rate = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Rate);
            if (!string.IsNullOrEmpty(data.ItemArray[34].ToString()))
                Actions.SetText(rate, data.ItemArray[34].ToString());
        }

        public static void EnterAgencySearchData(DataRow data)
        {
            var createOrderWindow = GetCreateNewOrderWindowProperties();

            var agencyCountry = Actions.GetWindowChild(createOrderWindow, NewOrderControls.Country);
            if (!string.IsNullOrEmpty(data.ItemArray[18].ToString()))
                DropDownActions.SelectDropdownByText(agencyCountry, data.ItemArray[18].ToString());

            var agencyState = Actions.GetWindowChild(createOrderWindow, NewOrderControls.State);
            if (!string.IsNullOrEmpty(data.ItemArray[19].ToString()))
                DropDownActions.SelectDropdownByText(agencyState, data.ItemArray[19].ToString());
        }

        public static void CloseCreateNewOrderWindow()
        {
            var createOrderWindow = GetCreateNewOrderWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)createOrderWindow);
            createOrderWindow.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        #endregion

        #region Search Order Window Methods

        public static void EnterDataInSearchOrderWindow(DataRow data)
        {
            var searchWindow = GetSearchOrderWindowProperties();

            var ssn = Actions.GetWindowChild(searchWindow, SearchOrderControls.Ssn);
            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
                Actions.SetText(ssn, data.ItemArray[3].ToString());

            var firstName = Actions.GetWindowChild(searchWindow, SearchOrderControls.FirstName);
            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
                Actions.SetText(firstName, data.ItemArray[4].ToString());

            var lastName = Actions.GetWindowChild(searchWindow, SearchOrderControls.LastName);
            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                Actions.SetText(lastName, data.ItemArray[5].ToString());

            var agencyNo = Actions.GetWindowChild(searchWindow, SearchOrderControls.AgencyNo);
            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
                Actions.SetText(agencyNo, data.ItemArray[6].ToString());

            var agencyName = Actions.GetWindowChild(searchWindow, SearchOrderControls.AgencyName);
            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
                Actions.SetText(agencyName, data.ItemArray[7].ToString());

            var primaryCaseNo = Actions.GetWindowChild(searchWindow, SearchOrderControls.PrimaryCaseNo);
            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
                Actions.SetText(primaryCaseNo, data.ItemArray[8].ToString());

            var secondaryCaseNo = Actions.GetWindowChild(searchWindow, SearchOrderControls.SecondaryCaseNo);
            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
                Actions.SetText(secondaryCaseNo, data.ItemArray[9].ToString());

            MouseActions.ClickButton(searchWindow, SearchOrderControls.SearchBtn);
            Playback.Wait(2000);

            var orderStatus = Actions.GetWindowChild(searchWindow, SearchOrderControls.OrderStatus);
            if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
                DropDownActions.SelectDropdownByText(orderStatus, data.ItemArray[10].ToString());
            Playback.Wait(2000);

            TableActions.OpenRecordFromTable(searchWindow, SearchOrderControls.ResultsGrid, "Social Security Number",
                data.ItemArray[11].ToString());
        }

        public static bool VerifySearchOrderWindowDisplayed()
        {
            var searchOrder = GetSearchOrderWindowProperties();
            return searchOrder.Exists;
        }

        public static void ClickOnAddNotesBtn()
        {
            var searchOrder = GetSearchOrderWindowProperties();
            MouseActions.ClickButton(searchOrder, "btnAddNotes");
        }

        public static void CloseSearchOrderWindow()
        {
            var searchOrder = GetSearchOrderWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)searchOrder);

            searchOrder.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        public static void ClickCloseBtnSearchOrderWindow()
        {
            var searchOrder = GetSearchOrderWindowProperties();
            MouseActions.ClickButton(searchOrder, "btnClose");
        }


        #endregion

        #region Order Details Window Methods

        public static bool VerifyOrderDetailsWindowDisplayed()
        {
            var orderDetails = GetOrderDetailsWindowProperties();
            return orderDetails.Exists;
        }

        private static UITestControl GetTabs()
        {
            var orderDetails = GetOrderDetailsWindowProperties();
            var children = orderDetails.GetChildren();
            return children[3];
        }

        public static void SelectNotesTab()
        {
            var tabs = GetTabs();
            var tabList = tabs.Container.SearchFor<WinWindow>(new { ControlName = "orderSummaryWorkspace" });
            var selectedTab = tabList.Container.SearchFor<WinTabPage>(new { Name = "Notes" });
            MouseActions.Click(selectedTab);

        }

        public static void CloseOrderDetailsWindow()
        {
            var orderDetails = GetOrderDetailsWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)orderDetails);
            orderDetails.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }



        #endregion

        #region Add Note Window Methods

        public static void EnterDatainAddNoteWindow(DataRow data)
        {
            var addnoteWindow = GetAddNoteWindowProperties();

            var addNote = Actions.GetWindowChild(addnoteWindow, "txtNotes");
            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
                Actions.SetText(addNote, data.ItemArray[12].ToString());

            MouseActions.ClickButton(addnoteWindow, "btnSaveNote");
        }

        public static bool VerifyAddNOteWindowDisplayed()
        {
            var addNoteWindow = GetAddNoteWindowProperties();
            return addNoteWindow.Exists;
        }

        #endregion

        #region Search Agency Window Methods

        public static void SelectAgencyRecord(DataRow data)
        {
            var searchAgency = GetSearchAgencyWindowProperties();

            TableActions.OpenRecordFromTable(searchAgency, "grdSearchAgency", "Agency No", data.ItemArray[35].ToString());
        }

        public static bool VerifySearchAgencyWindowDisplayed()
        {
            var searchAgency = GetSearchAgencyWindowProperties();
            return searchAgency.Exists;
        }

        public static void CloseSearchAgencyWindow()
        {
            var searchAgency = GetSearchAgencyWindowProperties();

            //TitlebarActions.ClickClose((WinWindow)searchAgency);

            searchAgency.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        #endregion

        #region Update Agency Window Methods

        public static bool VerifyUpdateAgencyWindowDisplayed()
        {
            var updateagency = GetUpdateAgencyWindowProperties();
            return updateagency.Exists;
        }

        public static void CloseUpdateAgencyWindow()
        {
            var updateagency = GetUpdateAgencyWindowProperties();

            //TitlebarActions.ClickClose((WinWindow)updateagency);

            updateagency.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        #endregion

        #region Answer Letter Window Methods

        public static bool VerifyAnswerLetterWindowDisplayed()
        {
            var answerLetter = GetAnswerLetterWindowProperties();
            return answerLetter.Exists;
        }

        public static void CloseAnswerLetterWindow()
        {
            var answerLetter = GetAnswerLetterWindowProperties();

            //TitlebarActions.ClickClose((WinWindow)answerLetter);

            answerLetter.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        public static void ClickonPrintBtnAnswerLetter()
        {
            var answerWindow = GetAnswerLetterWindowProperties();
            MouseActions.ClickButton(answerWindow, "btnPrint");
        }

        public static void ClickonSelectAllBtnAnswerLetter()
        {
            var answerWindow = GetAnswerLetterWindowProperties();
            MouseActions.ClickButton(answerWindow, "btnSelectAll");
        }

        public static bool ClosePrintWindow()
        {
            var printWindow = GetprintWindowProperties();
            if (printWindow.Exists)
            {
                printWindow.SetFocus();
                SendKeys.SendWait("{ESC}");
                Playback.Wait(2000);
                return true;
            }
            return false;
        }

        public static bool VerifyPrintWindowDisplayed()
        {
            var printWindow = GetprintWindowProperties();
            return printWindow.Exists;
        }

        #endregion

        #region Rate Sheet Window Methods

        public static bool VerifyRateSheetWindowDisplayed()
        {
            var rateSheet = GetRateSheetWindowProperties();
            return rateSheet.Exists;
        }

        public static void CloseRateSheetWindow()
        {
            var rateSheet = GetRateSheetWindowProperties();

            //TitlebarActions.ClickClose((WinWindow)rateSheet);
            rateSheet.SetFocus();
            SendKeys.SendWait("%s{F4}");
            Playback.Wait(2000);
        }

        public static void EnterRateSheetData(DataRow data)
        {
            var rateSheet = GetRateSheetWindowProperties();

            var cell = TableActions.SelectCellFromTable(rateSheet, "grdRateSheet", "RateSheetDetailsDomain row 1",
                "RateSheetDate");
            cell.SetFocus();
            Playback.Wait(2000);
            Mouse.Click(cell);
            SendKeys.SendWait("{HOME}");
            SendKeys.SendWait(data.ItemArray[36].ToString());
            Playback.Wait(2000);

            MouseActions.ClickButton(rateSheet, "btnUpdate");
        }

        #endregion

        #region Rate Sheet Date Window Methods

        public static bool VerifyRateSheetDateWindowDisplayed()
        {
            var rateSheetDate = GetRateSheetDateWindowProperties();
            return rateSheetDate.Exists;
        }

        public static void ClickOkinRateSheetDateWindow()
        {
            var rateSheetDate = GetRateSheetDateWindowProperties();
            MouseActions.ClickButton(rateSheetDate, "ButtonAccept");
        }

        #endregion

        #region Prevailing Wage Job Methods

        public static bool VerifySummaryViewTitleValue(string prevailingWageJob)
        {

            var summaryViewTitle = Actions.GetWindowChild(EllisWindow, "SummaryViewTitle");
            if (summaryViewTitle.Exists)
            {
                var summaryName = summaryViewTitle.GetProperty("Name");
                if (summaryName.Equals(prevailingWageJob))
                    return true;
            }
            return false;

        }

        #endregion

        #region Overtime Window Methods

        public static void EnterDatainOverTimeWindow(DataRow data)
        {

            var pane = EllisWindow.Container.SearchFor<WinComboBox>(new { Name = "" });
            var ddConCol = Actions.GetControlCollection(pane);
            var country = ddConCol.Where(v => v.Enabled && v.GetProperty("ControlName").Equals("_ddlCountry")).First();
            country.SetFocus();
            SendKeys.SendWait("UNITED STATES");
            Playback.Wait(1000);
            var state = ddConCol.Where(v => v.Enabled && v.GetProperty("ControlName").Equals("_ddlState")).First();
            state.SetFocus();
            if (data.ItemArray[8].ToString().CompareTo(String.Empty) != 0)
            {
                SendKeys.SendWait(data.ItemArray[8].ToString());
            }
            else
            {
                SendKeys.SendWait("ALL");
            }

            Playback.Wait(1000);

            MouseActions.ClickButton(EllisWindow, "_btnSearch");
            Playback.Wait(2000);
            TableActions.OpenRecordFromTable(EllisWindow, "grdOvertime", "State", data.ItemArray[11].ToString());

        }

        public static bool VerifyOverTimeWindowDisplayed()
        {
            var overTime = GetOverTimeWindowProperties();
            return overTime.Exists;
        }

        public static void CloseOverTimeWindow()
        {
            var overTime = GetOverTimeWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)overTime);
            overTime.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        public static void ClickModifyBtnOverTime()
        {
            MouseActions.ClickButton(GetOverTimeWindowProperties(), "btnModify");
        }

        public static void ClickokBtnSaveSucccess()
        {
            MouseActions.ClickButton(GetSaveSuccessWindowProperties(), "_OKButton");
        }



        #endregion

        #region State Holiday Window Methods

        public static void EnterDataInStateHoliday(DataRow data)
        {
            var pane = EllisWindow.Container.SearchFor<WinComboBox>(new { Name = "" });
            var ddConCol = Actions.GetControlCollection(pane);
            foreach (var control in ddConCol)
            {
                if (control.Enabled && control.GetProperty("ControlName").Equals("ddlCountry"))
                {
                    control.SetFocus();
                    SendKeys.SendWait("UNITED STATES");
                    Playback.Wait(1000);
                }
                if (control.Enabled && control.GetProperty("ControlName").Equals("ddlState"))
                {
                    control.SetFocus();
                    SendKeys.SendWait(data.ItemArray[8].ToString());
                }
                if (control.Enabled && control.GetProperty("ControlName").Equals("ddlYear"))
                {
                    control.SetFocus();
                    SendKeys.SendWait(data.ItemArray[9].ToString());
                }
            }

            MouseActions.ClickButton(EllisWindow, "btnSearch");
            Playback.Wait(2000);

        }


        #endregion

        #region No Results Found Window Methods

        public static bool VerifyNoResultsFoundWindowDisplayed()
        {
            var noWindow = GetNoResultsFoundWindowProperties();
            return noWindow.Exists;
        }

        public static void ClickOnOkBtnNoresults()
        {
            MouseActions.ClickButton(GetNoResultsFoundWindowProperties(), "ButtonAccept");
        }
        #endregion

        #region Payment Refund Window Methods

        public static void EnterDataInPaymentRefundStatusWindow(DataRow data)
        {
            var refundStatus = Actions.GetWindowChild(EllisWindow, "cmbRefundStatus");
            Factory.AssertIsFalse(refundStatus.Equals(String.Empty),
                "The datarow contains an empty string where it should contain a value.");
            DropDownActions.SelectDropdownByText(refundStatus, data.ItemArray[10].ToString());

            MouseActions.ClickButton(EllisWindow, "btnGO");
        }


        #endregion

        #region Search Payment Methods

        public static void CloseSearchPaymentsWindow()
        {
            var closeButton = GetPaymentSearchCloseWindowButton();
            Mouse.Click(closeButton.GetClickablePoint());
            
        }

        public static bool VerifySearchPaymentsWindowDisplayed()
        {
            var searchPayment = GetSearchPaymentsWindowProperties();
            return searchPayment.Exists;
        }

        public static void ClickOnResetBtnPaymentSearch()
        {
            var searchPayment = GetSearchPaymentsWindowProperties();
            MouseActions.ClickButton(searchPayment, "btnNewSearch");
        }

        public static void ClickOnSearchBtnPaymentSearch()
        {
            var searchPayment = GetSearchPaymentsWindowProperties();
            MouseActions.ClickButton(searchPayment, "btnSearch");
        }

        public static void EnterDataInSearchPayment(DataRow data)
        {
            var searchPayment = GetSearchPaymentsWindowProperties();

            if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
            {
                var ssn = Actions.GetWindowChild(searchPayment, "mskSocialSecurityNumber");
                Actions.GetWindowChild(ssn, data.ItemArray[3].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
            {
                var firstName = Actions.GetWindowChild(searchPayment, "txtFirstName");
                Actions.GetWindowChild(firstName, data.ItemArray[4].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
            {
                var lastName = Actions.GetWindowChild(searchPayment, "txtLastName");
                Actions.GetWindowChild(lastName, data.ItemArray[5].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
            {
                var agencyNo = Actions.GetWindowChild(searchPayment, "txtAgencyNumber");
                Actions.GetWindowChild(agencyNo, data.ItemArray[6].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
            {
                var apVendor = Actions.GetWindowChild(searchPayment, "txtVendorNumber");
                Actions.GetWindowChild(apVendor, data.ItemArray[7].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[8].ToString()))
            {
                var primaryCase = Actions.GetWindowChild(searchPayment, "txtPrimaryCase");
                Actions.GetWindowChild(primaryCase, data.ItemArray[8].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[9].ToString()))
            {
                var secondaryCase = Actions.GetWindowChild(searchPayment, "txtSecondaryCase");
                Actions.GetWindowChild(secondaryCase, data.ItemArray[9].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[10].ToString()))
            {
                var paymentMethod = Actions.GetWindowChild(searchPayment, "cmbPaymentMethod");
                DropDownActions.SelectDropdownByText(paymentMethod, data.ItemArray[10].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[11].ToString()))
            {
                var typeOfOrder = Actions.GetWindowChild(searchPayment, "cmbTypeOfOrder");
                DropDownActions.SelectDropdownByText(typeOfOrder, data.ItemArray[11].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[12].ToString()))
            {
                //var fromDate = Actions.GetWindowChild(searchPayment, "dtFromDateRange");
                #region HackToOvercomeNoNameOnDateRangeStart
                Actions.GetWindowChild(searchPayment, "cmbTypeOfOrder").SetFocus();
                Keyboard.SendKeys("{TAB}");

                for (int i = 0; i < 10; i++)
                {
                    System.Threading.Thread.Sleep(35);
                    Keyboard.SendKeys("+{RIGHT}");
                }

                Keyboard.SendKeys("{DEL}");

                var arr = data.ItemArray[12].ToString().Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var datePart in arr)
                {
                    Keyboard.SendKeys(datePart);
                    Keyboard.SendKeys("{RIGHT}");
                } 
                #endregion
                //DropDownActions.SelectDropdownByText(fromDate, data.ItemArray[12].ToString());
            }

            if (!string.IsNullOrEmpty(data.ItemArray[13].ToString()))
            {
                var toDate = Actions.GetWindowChild(searchPayment, "dtToDateRange");
                DropDownActions.SelectDropdownByText(toDate, data.ItemArray[13].ToString());
            }
        }

        #endregion

        #region Cleared Check Gateway Methods

        public static void CloseClearedCheckGatewayWindow()
        {
            var gateWay = GetClearedCheckGatewayWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)gateWay);
            gateWay.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        public static bool VerifyClearedCheckGatewayWindowDisplayed()
        {
            var gateWay = GetClearedCheckGatewayWindowProperties();
            return gateWay.Exists;
        }

        #endregion

        #region Schedule Payments Window Methods

        public static void CloseSchedulePaymentWindow()
        {
            var schedule = GetSchedulePaymentsWindowProperties();
            //TitlebarActions.ClickClose((WinWindow)schedule);

            schedule.SetFocus();
            SendKeys.SendWait("%{F4}");
            Playback.Wait(2000);
        }

        public static bool VerifySchedulePaymentWindowDisplayed()
        {
            var schedule = GetSchedulePaymentsWindowProperties();
            return schedule.Exists;
        }


        #endregion

        #region Contorls

        private class NewOrderControls
        {
            public const string Recipient = "ddlRecipient";
            public const string Ssn = "txtMaskSSN";
            public const string FirstName = "txtFirstName";
            public const string MiddleName = "txtMName";
            public const string LastName = "txtLastName";
            public const string BirthDate = "dtBirthDate";
            public const string FirstWorkDt = "txtFirstDate";
            public const string SearchWorkerBtn = "btnSearchWorker";
            public const string WorkerStatus = "txtStatus";
            public const string AppliedForWork = "txtAppliedForWork";
            public const string GenerateLetterBtn = "btnGenerateLetter";
            public const string MartialStatus = "txtMaritalStatus";
            public const string LastWorkDt = "txtLastWorkDate";
            public const string AgencyNo = "txtAgencyNo";
            public const string ApVendor = "txtAPVendor";
            public const string AgencyName = "txtAgencyName";
            public const string Zip = "txtAZip";
            public const string Country = "ddlCountry";
            public const string State = "ddlState";
            public const string SearchAgencyBtn = "btnSearchAgency";
            public const string TypeOfOrder = "ddlOrderType";
            public const string ViewOrderImageBtn = "btnViewOimage";
            public const string PrimaryCaseNo = "txtPrimaryCaseNo";
            public const string OrderDetailsBtn = "btnODetails";
            public const string SecondaryCaseNo = "txtSecCaseNo";
            public const string OriginalOrderDt = "dtOriginalOrderDate";
            public const string SearchOrderBtn = "btnSearch";
            public const string OrderStatus = "ddlStatus";
            public const string ReleaseVerify = "ddlReleaseverify";
            public const string StatusOrderDt = "dtStatusOrderDate";
            public const string Arrears = "ddlArrears";
            public const string ExpirationMethodDt = "ddlExpiration";
            public const string ExpiredDt = "dtExpiredDate";
            public const string OrderPriority = "textOrderPriority";
            public const string Deduct = "txtDeduct";
            public const string Per = "ddlPer";
            public const string TotalMax = "txtTotalMax";
            public const string Rate = "txtRate";
            public const string AcceptBtn = "btnAccept";
            public const string TreatAsNewBtn = "btnTreatasNew";
            public const string ResetBtn = "btnReset";
            public const string AttachImage = "btnAttachImage";

        }

        private class SearchOrderControls
        {
            public const string Ssn = "txtMaskSSN";
            public const string FirstName = "txtFirstName";
            public const string LastName = "txtLastName";
            public const string AgencyNo = "txtAgencyNumber";
            public const string AgencyName = "txtAgencyName";
            public const string PrimaryCaseNo = "txtPrimaryCAseNumber";
            public const string SecondaryCaseNo = "txtSecondaryCaseNumber";
            public const string NewSearchBtn = "btnNewSearch";
            public const string SearchBtn = "btnSearch";
            public const string CloseBtn = "btnClose";
            public const string OrderStatus = "ddlOrderStatus";
            public const string ResultsGrid = "grdOrderResults";

        }

        #endregion


    }
}
