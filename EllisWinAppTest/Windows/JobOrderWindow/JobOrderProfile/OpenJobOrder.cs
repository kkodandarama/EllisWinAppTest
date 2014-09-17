using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Windows.JobOrderWindow.JobOrderProfile
{
    internal class OpenJobOrder : AppContext
    {
        public static bool SelectJobOrderFromTable()
        {
            LandingPage.SelectFromToolbar("Job Orders");
            //Factory.GetAllControlNames(EllisWindow);
            var tableName = Actions.GetWindowChild(EllisWindow, "_grdJobOrders");
            var table = (WinTable)tableName;

            var row = table.Container.SearchFor<WinRow>(new { Name = "DispatchJobOrderDetailSummary row 1" });
            var cell = row.Container.SearchFor<WinCell>(new { Name = "Job Order #" });
            Globals.JobOrderNo = cell.Value;
            Mouse.DoubleClick(cell);


            var profileWindow = JobOrderProfileWindowProperties();
            return profileWindow.Exists;
        }

        public static bool OpenRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string JoNumber)
        {
            
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;
            
            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue == JoNumber)
                {
                    Mouse.Click(rowHeader);
                    Mouse.DoubleClick(rowHeader);
                    break;
                }

            }

            var row = table.Container.SearchFor<WinRow>(new { Name = "DispatchJobOrderDetailSummary row 1" });
            var cell = row.Container.SearchFor<WinCell>(new { Name = "Job Order #" });
            Globals.JobOrderNo = cell.Value;
            Mouse.DoubleClick(cell);

            var profileWindow = JobOrderProfileWindowProperties();
            return profileWindow.Exists;
        }

        public static bool SelectExpiredJobOrderFromTable()
        {
            var view = EllisWindow.Container.SearchFor<WinMenuItem>(new { Name = "Views" });
            var dispOrPout = view.Container.SearchFor<WinMenuItem>(new { Name = "Job Orders" });
            var window = dispOrPout.Container.SearchFor<WinMenuItem>(new { Name = "Expired" });

            MouseActions.Click(view);
            MouseActions.Click(dispOrPout);
            MouseActions.Click(window);

            //Factory.GetAllControlNames(EllisWindow);
            var tableName = Actions.GetWindowChild(EllisWindow, "_grdJobOrders");
            var table = (WinTable)tableName;

            var row = table.Container.SearchFor<WinRow>(new { Name = "RecentJobOrderSummary row 1" });
            if (row.Exists)
            {
                var cell = row.Container.SearchFor<WinCell>(new { Name = "Job Order #" });
                cell.SetFocus();
                Globals.JobOrderNo = cell.Value;
                Mouse.DoubleClick(cell);
            }
            var profileWindow = JobOrderProfileWindowProperties();
            return profileWindow.Exists;
        }

        //public static UITestControl JobOrderNoProfileWindowProperties()
        //{
        //    var joborderWindow = App.Container.SearchFor<WinWindow>(new { Name = Globals.JobOrderNo + "-Job Order Profile" });
        //    var winGroup = joborderWindow.Container.SearchFor<WinGroup>(new { Name = "" });
        //    return winGroup;
        //}

        public static UITestControl JobOrderProfileWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return joborderWindow;
        }

        public static void CloseJobOrderProfile()
        {
            var profileWindow = JobOrderProfileWindowProperties();
            if (profileWindow.Exists)
                MouseActions.ClickButton(profileWindow, "btnClose");
        }

        public static bool VerifyJobOrderWindowDisplayed(string data)
        {
            var window = JobOrderProfileWindowProperties();
            return window.Name.Contains(data);
        }

        public static bool VerifyJobOrder(DataRow data)
        {
            bool match = true;

            var profileWindow = JobOrderProfileWindowProperties();

            //Verify Data in Summary Tab
            //---------------------------------------------------------------------------------------------
            //JobOrderNumber
            //---------------------------------------------------------------------------------------------

            var lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobOrderNumber);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[2].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[2].ToString().Trim());
            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[2].ToString().Trim());
                match = false;
            }
            //---------------------------------------------------------------------------------------------
            //CustNumber
            //---------------------------------------------------------------------------------------------

            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.CustNumber);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[3].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[3].ToString().Trim());
            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[3].ToString().Trim());
                match = false;
            }

            //---------------------------------------------------------------------------------------------
            //QuoteNumber
            //---------------------------------------------------------------------------------------------


            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.QuoteNumber);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[4].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[4].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[4].ToString().Trim());
                match = false;

            }
            //---------------------------------------------------------------------------------------------
            //QuoteType
            //---------------------------------------------------------------------------------------------

            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.QuoteType);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[5].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[5].ToString().Trim());


            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[5].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //CompCodeDescription 
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.CompCodeDescription);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[6].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[6].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[6].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //PrevailingWage 
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.PrevailingWage);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[7].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[7].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[7].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //WrapInsurance 
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.WrapInsurance);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[8].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[8].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[8].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobSiteName 
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobSiteName);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[9].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[9].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[9].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobSiteAddress  
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobSiteAddress);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[10].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[10].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[10].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobSiteContact  
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobSiteContact);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[11].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[11].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[11].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //CustomerPhone  
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.CustomerPhone);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[12].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[12].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[12].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobDuties  
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobDuties);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[13].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[13].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[13].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //PayRate  
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.PayRate);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[14].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[14].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[14].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //BillRate   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.BillRate);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[15].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[15].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[15].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //SafetyEquipment   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.SafetyEquipment);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[16].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[16].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[16].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //IsECustomerOrder   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.IsECustomerOrder);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[17].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[17].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[17].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobOrderEffective   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobOrderEffective);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[18].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[18].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[18].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JoExpiration   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JoExpiration);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[19].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[19].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[19].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //RepeatsAllowed   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.RepeatsAllowed);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[20].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[20].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[20].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //PoNumber   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.PoNumber);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[21].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[21].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[21].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //JobOrderPaymentType   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.JobOrderPaymentType);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[22].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[22].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[22].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //OrderOrderPlacedBy   
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.OrderOrderPlacedBy);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[23].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[23].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[23].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //BranchNumber    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.BranchNumber);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[24].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[24].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[24].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //AssociateName    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.AssociateName);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[25].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[25].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[25].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //OrderDate    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.OrderDate);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[26].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[26].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[26].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //DrugTest    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.DrugTest);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[27].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[27].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[27].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //BgCheck    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.BgCheck);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[28].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[28].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[28].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //SpecialNotes    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.SpecialNotes);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[29].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[29].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[29].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //CertificationRequired     
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.CertificationRequired);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[30].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[30].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[30].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //LicenseRequired    
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.LicenseRequired);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[31].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[31].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[31].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //SkillsRequired     
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.SkillsRequired);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[32].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[32].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[32].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //RecentEvaluationInfo     
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.RecentEvaluationInfo);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[33].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[33].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[33].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //RecentEvaluationDate     
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.RecentEvaluationDate);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[34].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[34].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[34].ToString().Trim());
                match = false;

            }

            //---------------------------------------------------------------------------------------------
            //ScheduledEvaluationDate     
            //---------------------------------------------------------------------------------------------
            lblControl = Actions.GetWindowChild(profileWindow, JobOrderSummaryConstents.ScheduledEvaluationDate);
            if (lblControl.GetProperty("Name").ToString() == data.ItemArray[35].ToString().Trim())
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- matched with -->  " +
                                  data.ItemArray[35].ToString().Trim());



            }

            else
            {
                Console.WriteLine(lblControl.GetProperty("Name").ToString() + " <-- didn't match with -->  " +
                                  data.ItemArray[35].ToString().Trim());
                match = false;

            }


            return match;
        }


        public class JobOrderSummaryConstents
        {
            public const string JobOrderNumber = "lblJobOrderNumber";
            public const string CustNumber = "_lblCustomerName";
            public const string QuoteNumber = "_lblQuoteNumber";
            public const string QuoteType = "_lblQuoteType";
            public const string CompCodeDescription = "_lblCompCodeDescription";
            public const string PrevailingWage = "_lblPrevailingWage";
            public const string WrapInsurance = "_lblWrapInsurance";
            //  "Job Site";
            public const string JobSiteName = "_lblJobSiteName";
            public const string JobSiteAddress = "_lblJobSiteAddress";
            public const string JobSiteContact = "_lblJobSiteContact";
            public const string CustomerPhone = "_lblCustomerPhone";
            //  "Job Duty";
            public const string JobDuties = "_lblJobDuties";
            public const string PayRate = "_lblPayRate";
            public const string BillRate = "_lblBillRate";
            //  "Safety Eqe";
            public const string SafetyEquipment = "_lblSafetyEquipment";
            //  "Job order";
            public const string IsECustomerOrder = "_lblIsECustomerOrder";
            public const string JobOrderEffective = "_lblJobOrderEffective";
            public const string JoExpiration = "_lblJOExpiration";
            public const string RepeatsAllowed = "_lblRepeatsAllowed";
            public const string PoNumber = "_lblPONumber";
            public const string JobOrderPaymentType = "_lblJobOrderPaymentType";
            public const string OrderOrderPlacedBy = "_lblOrderOrderPlacedBy";
            public const string BranchNumber = "_lblBranchNumber";
            public const string AssociateName = "_lblAssociateName";
            public const string OrderDate = "_lblOrderDate";
            //  "Worker Requirements";
            public const string DrugTest = "_lblDrugTest";
            public const string BgCheck = "_lblBGCheck";
            public const string SpecialNotes = "_lblSpecialNotes";
            public const string CertificationRequired = "_lblCertificationRequired";
            public const string LicenseRequired = "_lblLicenseRequired";
            public const string SkillsRequired = "_lblSkillsRequired";
            //  "Site Safety";
            public const string RecentEvaluationInfo = "_lblRecentEvaluationInfo";
            public const string RecentEvaluationDate = "_lblRecentEvaluationDate";
            public const string ScheduledEvaluationDate = "_lblScheduledEvaluationDate";
            // "Other Requirements";
            public const string CommonReq = "_lblCommonReq";



        }

        public class JobOrderBasicJobOrderConstants
        {
            // Editbox control constants
            public const string AddressLine1 = "txtAddressLine1";
            public const string AddressLine2 = "txtAddressLine2";
            public const string BillingAddressLine1 = "txtBillingAddressLine1";
            public const string BillingAddressLine2 = "txtBillingAddressLine2";
            public const string BillingEmail = "txtBillingEmail";
            public const string BillingFax = "txtBillingFax";
            public const string BillingPhone = "txtBillingPhone";
            public const string BillingZip = "txtBillingZip";
            public const string BillRate = "mskBillRate";
            public const string City = "txtCity";
            public const string CustomerCostCentre = "txtCustomerCostCentre";
            public const string CustomerProjectName = "txtCustomerProjectName";
            public const string GrossMargin = "mskGrossMargin";
            public const string JobDuty = "txtJobDuty";
            public const string JobSiteEmail = "txtJobSiteEmail";
            public const string JobSiteFax = "txtJobSiteFax";
            public const string JobSiteName = "txtJobSiteName";
            public const string JobSiteNewContact = "txtJobSiteNewContact";
            public const string JobSitePhone = "txtJobSitePhone";
            public const string JobSiteZip = "txtJobSiteZip";
            public const string LastEvaluationDate = "txtLastEvaluationDate";
            public const string OrderPlacedByContact = "txtOrderPlacedByContact";
            public const string PayRate = "txtPayRate";
            public const string PhoneNumber = "txtPhoneNumber";
            public const string PoNumber = "txtPONumber";
            public const string ReportToAddressLine1 = "txtReportToAddressLine1";
            public const string ReportToAddressLine2 = "txtReportToAddressLine2";
            public const string ReportToEmail = "txtReportToEmail";
            public const string ReportToFax = "txtReportToFax";
            public const string ReportToPhone = "txtReportToPhone";
            public const string ReportToZip = "txtReportToZip";


            // ComboBox control constants

            public const string JobSiteNameComboBox = "cmbJobSiteName";
            public const string JobSiteContact = "cmbJobSiteContact";
            public const string ReportingContacts = "cmbReportingContacts";
            public const string EffectiveDate = "dtpEffectiveDate";
            public const string ExpirationDate = "dtpExpirationDate";
            public const string State = "cmbState";


        }

        public static bool EditJobOrder(DataRow data)
        {
            var profileStatus = true;
            var window = JobOrderProfileWindowProperties();
            if (window.Exists)
            {

            }

            return profileStatus;
        }

        private static UITestControl GetTabs()
        {
            var joProfileWindow = JobOrderProfileWindowProperties();
            var children = joProfileWindow.GetChildren();
            return children[3];
        }

        public static void SelectTab(string tabName)
        {
            var tabs = GetTabs();
            var tabList = tabs.Container.SearchFor<WinWindow>(new { ControlName = "TabJobOdrer" });
            var selectedTab = tabList.Container.SearchFor<WinTabPage>(new { Name = tabName });
            MouseActions.Click(selectedTab);
        }

        internal static void EditBasicJobInfoOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {

                // -----------------------------------------------------------------------------
                // AddressLine1
                // -----------------------------------------------------------------------------

                var controlName = JobOrderBasicJobOrderConstants.AddressLine1;
                var control = Actions.GetWindowChild(windowInst, controlName);
                var controlData = data.ItemArray[3].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // AddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.AddressLine2;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[4].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingAddressLine1
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingAddressLine1;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[5].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingAddressLine2;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[6].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingEmail
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingEmail;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[7].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingFax
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingFax;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[8].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingPhone
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingPhone;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[9].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }

                // -----------------------------------------------------------------------------
                // BillingZip
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillingZip;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[10].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillRate
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.BillRate;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[11].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // City
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.City;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[12].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.CustomerCostCentre;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[13].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.CustomerProjectName;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[14].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.GrossMargin;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[15].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobDuty;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[16].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteEmail;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[17].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteFax;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[18].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteName;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[19].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteNewContact;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[20].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSitePhone;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[21].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteZip;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[22].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.LastEvaluationDate;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[23].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.OrderPlacedByContact;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[24].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.PayRate;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[25].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.PhoneNumber;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[26].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.PoNumber;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[27].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToAddressLine1;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[28].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToAddressLine2;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[29].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToEmail;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[30].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToFax;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[31].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToPhone;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[32].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportToZip;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[33].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    Actions.SetText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteName;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[34].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.JobSiteContact;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[35].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ReportingContacts;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[36].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.EffectiveDate;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[37].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.ExpirationDate;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[38].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }
                // -----------------------------------------------------------------------------
                // BillingAddressLine2
                // -----------------------------------------------------------------------------

                controlName = JobOrderBasicJobOrderConstants.State;
                control = Actions.GetWindowChild(windowInst, controlName);
                controlData = data.ItemArray[39].ToString();
                if (controlData != "" && control.Enabled.Equals(true))
                {
                    DropDownActions.SelectDropdownByText(control, controlData);
                }

                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
                //Factory.GetAllControlNames(windowInst);
            }
        }

        public static void EditOrderDetailsAddlChargesOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                //Enter Job Order Schedule Data
                JobOrderSchedule(data);

                //Enter Job Order Notes
                if (data.ItemArray[71].ToString() != "")
                {
                    MouseActions.ClickButton(windowInst, "");
                    var notesWindow = App.Container.SearchFor<WinWindow>(new { name = "Order Notes" });

                    var textOrderNotes = Actions.GetWindowChild(notesWindow, "_txtOrderNotes");
                    Actions.SetText(textOrderNotes, data.ItemArray[70].ToString());

                    var dropDownOrderNotes = Actions.GetWindowChild(notesWindow, "_cmbRequestedBy");
                    Playback.Wait(1000);
                    DropDownActions.SelectDropdownByText(dropDownOrderNotes, data.ItemArray[71].ToString());
                    MouseActions.ClickButton(notesWindow, "_btnOk");
                }

                // Save Job Order Details Additional Charges Info
                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }




        }

        public static void EditRequirementsOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                // Add Skills Data
                MouseActions.ClickButton(windowInst, "btnAddSkills");
                var addSkillsWindow = App.Container.SearchFor<WinWindow>(new { Name = "Add Skills" });
                if (addSkillsWindow.Exists)
                {
                    if (data.ItemArray[80].ToString() != "")
                    {
                        var posFocus = Actions.GetWindowChild(addSkillsWindow, "cmbPositionFocus");
                        DropDownActions.SelectDropdownByText(posFocus, data.ItemArray[79].ToString());

                        var posTitle = Actions.GetWindowChild(addSkillsWindow, "txtSearchText");
                        Actions.SetText(posTitle, data.ItemArray[80].ToString());

                        MouseActions.ClickButton(addSkillsWindow, "btnSearch");

                        var chkBox = Actions.GetWindowChild(addSkillsWindow, "chkSelect");
                        Mouse.Click(chkBox);

                        MouseActions.ClickButton(addSkillsWindow, "btnAdd");
                        MouseActions.ClickButton(addSkillsWindow, "btnAddSkillsExperience");
                    }
                    else
                    {
                        TitlebarActions.ClickClose(addSkillsWindow);
                    }

                }

                // Add Certificates

                if (data.ItemArray[75].ToString() != "")
                {
                    var tabRowCertifications = TableActions.SelectCellFromTable(windowInst, "grdCertifications",
                        "Template Add Row", "Certificate Type");
                    MouseActions.DoubleClick(tabRowCertifications);
                    SendKeys.SendWait(data.ItemArray[75].ToString());
                    SendKeys.SendWait("{TAB}");
                }

                // Add Lisenses
                if (data.ItemArray[76].ToString() != "")
                {
                    var tabRowLicense = TableActions.SelectCellFromTable(windowInst, "grdLicense", "Template Add Row",
                        "License Type");
                    MouseActions.DoubleClick(tabRowLicense);
                    SendKeys.SendWait(data.ItemArray[76].ToString());
                    SendKeys.SendWait("{TAB}");
                }

                // Add Background Check
                if (data.ItemArray[77].ToString() != "")
                {
                    var tabRowBackgroundCheck = TableActions.SelectCellFromTable(windowInst, "grdBackgroundCheck",
                        "Template Add Row", "Background Check Type");
                    MouseActions.DoubleClick(tabRowBackgroundCheck);
                    SendKeys.SendWait(data.ItemArray[77].ToString());
                    SendKeys.SendWait("{TAB}");
                }

                // Add Safety Requirements
                if (data.ItemArray[78].ToString() != "")
                {
                    var tabRowSafetyEquipment = TableActions.SelectCellFromTable(windowInst, "grdSafetyEquipment",
                        "Template Add Row", "Type");
                    MouseActions.DoubleClick(tabRowSafetyEquipment);
                    SendKeys.SendWait(data.ItemArray[78].ToString());
                    SendKeys.SendWait("{TAB}");
                }

                // Add Dress Code
                if (data.ItemArray[72].ToString() != "")
                {
                    var dressCode = Actions.GetWindowChild(windowInst, "txtDressCode");
                    Actions.SetText(dressCode, data.ItemArray[72].ToString());
                }

                // Add Special Notes
                if (data.ItemArray[73].ToString() != "")
                {
                    var splNotes = Actions.GetWindowChild(windowInst, "txtSpecialNotes");
                    Actions.SetText(splNotes, data.ItemArray[73].ToString());
                }
                // Add Common Requirements
                if (data.ItemArray[74].ToString() != "")
                {
                    var commonRequirements = Actions.GetWindowChild(windowInst, "txtCommonRequirements");
                    Actions.SetText(commonRequirements, data.ItemArray[74].ToString());
                }

                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }
        }

        public static void EditPreQualifyingQuestionsOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }
        }

        public static void EditSafetyOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                if (data.ItemArray[82].ToString() == "Yes" || data.ItemArray[82].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optWrittenSafetyProgram", data.ItemArray[82].ToString());
                if (data.ItemArray[83].ToString() == "Yes" || data.ItemArray[83].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optCustomerSafetyProgram", data.ItemArray[83].ToString());
                if (data.ItemArray[84].ToString() == "Yes" || data.ItemArray[84].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optPPE", data.ItemArray[84].ToString());
                if (data.ItemArray[85].ToString() == "Yes" || data.ItemArray[85].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optSpecialisedTraining", data.ItemArray[85].ToString());
                if (data.ItemArray[86].ToString() == "Yes" || data.ItemArray[86].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optElevatedSurface", data.ItemArray[86].ToString());
                if (data.ItemArray[87].ToString() == "Yes" || data.ItemArray[87].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optHazardousMaterials", data.ItemArray[87].ToString());
                if (data.ItemArray[88].ToString() == "Yes" || data.ItemArray[88].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optTrenchWork", data.ItemArray[88].ToString());
                if (data.ItemArray[89].ToString() == "Yes" || data.ItemArray[89].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optConfinedSpace", data.ItemArray[89].ToString());
                if (data.ItemArray[90].ToString() == "Yes" || data.ItemArray[90].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, "_optMovingMachinery", data.ItemArray[90].ToString());
                if (data.ItemArray[91].ToString() == "Yes" || data.ItemArray[91].ToString()=="No")
                    Factory.SelectRadioButton(windowInst, "_optWorkersOnsite", data.ItemArray[91].ToString());

                var winControl = Actions.GetWindowChild(windowInst, "_dtpSafetyEvaluationDate");
                if (winControl.Enabled)
                {
                   Factory.SetMaskedText(windowInst, "_dtpSafetyEvaluationDate", data.ItemArray[95].ToString());
                }


                if (data.ItemArray[92].ToString() != "")
                {
                    winControl = Actions.GetWindowChild(windowInst, "_txtComments");
                    Actions.SetText(winControl, data.ItemArray[92].ToString());
                }
                if (data.ItemArray[93].ToString() != "")
                {
                    winControl = Actions.GetWindowChild(windowInst, "_txtActionTaken");
                    Actions.SetText(winControl, data.ItemArray[93].ToString());
                }
                if (data.ItemArray[94].ToString() != "")
                {
                    winControl = Actions.GetWindowChild(windowInst, "_txtIdentifiedHazards");
                    Actions.SetText(winControl, data.ItemArray[94].ToString());
                }

                if (data.ItemArray[95].ToString() != "")
                    Factory.SetMaskedText(windowInst, "_dtpSiteEvaluated", data.ItemArray[95].ToString());
                
                if (data.ItemArray[96].ToString() != "")
                {
                    var siteEvaluated =
                        windowInst.Container.SearchFor<WinRadioButton>(new { Name = "Satisfactory" });
                    Actions.SelectRadioButton(siteEvaluated);
                }

                MouseActions.ClickButton(windowInst, "btnAdd");
                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }
        }


        public static void EditProgressBillingOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                MouseActions.ClickButton(windowInst, "btnSave");
                CustomerWindow.CustomerProfileWindow.CloseWarningWindow();
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }
        }

        public static void JobOrderSchedule(DataRow data)
        {
            var winInst = JobOrderProfileWindowProperties();
            var tableName = Actions.GetWindowChild(winInst, "_grdJobOrderSchedule");
            var table = (WinTable)tableName;
            //var tablecont = table.Rows.GetNamesOfControls();

            var row = table.Container.SearchFor<WinRow>(new { Name = "Template Add Row" });

            var cell = row.Container.SearchFor<WinCell>(new { Name = "Start Time" });
            Mouse.DoubleClick(cell);
            SendKeys.SendWait("0905 PM");


            var cell2 = row.Container.SearchFor<WinCell>(new { Name = "Assigned To Branch" });
            Mouse.Click(cell2);
            SendKeys.SendWait("1711");

            //var cell3 = row.Container.SearchFor<WinCell>(new { Name = "MAP" });
            //Mouse.Click(cell3);

            var weekly = Actions.GetWindowChild(winInst, "_txtWeeklyDate");
            var weeklyData = weekly.GetProperty("Value").ToString();
            weeklyData = weeklyData + " end";
            var weekStart = Actions.GetTextBetween(weeklyData, "of", "thru");
            var weekEnd = Actions.GetTextBetween(weeklyData, "thru", "end");
            DateTime weekStartDate = Convert.ToDateTime(weekStart);
            DateTime weekEndDate = Convert.ToDateTime(weekEnd);
            var daysLeftInWeek = weekEndDate - DateTime.Now;
            char[] chArray = daysLeftInWeek.ToString().ToCharArray();
            var days = chArray[0].ToString();

            var cell4 = row.Container.SearchFor<WinCell>(new { Name = "Repeat Dispatch Allowed?" });
            cell4.SetFocus();
            //if (cell4.GetProperty("Value").ToString() != data.ItemArray[55].ToString())
            //{
            //    Mouse.Click(cell4);
            //}


            switch (days)
            {
                case "6":
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[57].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[59].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[61].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[63].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                case "5":
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[59].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[61].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[63].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                case "4":

                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[61].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[63].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                case "3":

                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[63].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                case "2":
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[65].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                case "1":
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[67].ToString());
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");

                    break;

                case "0":
                    SendKeys.SendWait("{TAB}");
                    SendKeys.SendWait(data.ItemArray[69].ToString());
                    SendKeys.SendWait("{TAB}");
                    break;

                default:

                    if (days == "-")
                    {
                        Console.WriteLine("Schedule not updated since the date range is in past");
                    }
                    else
                    {
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[57].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[59].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[61].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[63].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[65].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[67].ToString());
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait("{TAB}");
                        SendKeys.SendWait(data.ItemArray[69].ToString());
                        SendKeys.SendWait("{TAB}");
                    }
                    break;


            }


            //cell = row.Container.SearchFor<WinCell>(new { Name = "Assigned to Branch" });
            //Mouse.DoubleClick(cell);


        }
    }
}