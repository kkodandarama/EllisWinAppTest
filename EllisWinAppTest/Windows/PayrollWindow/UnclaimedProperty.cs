using System;
using System.Data;
using System.Runtime.Remoting;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EllisWinAppTest.Windows.PayrollWindow
{
   public class UnclaimedProperty : AppContext
   {

       #region Window Properties

       public static UITestControl GetUnclaimedPropertyWindowProperties()
       {
           return Actions.GetWindowProperties(App, "Unclaimed Property");
       }

       private static UITestControl GetprintWindowProperties()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           var printWindow = unclaimedProperty.Container.SearchFor<WinWindow>(new { Name = "Print" });
           return printWindow;
       }

       #endregion

       #region Unclaimed Property Print Letter Window Methods

       public static bool VerifyUnclaimedPropertyWindowDisplayed()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           return unclaimedProperty.Exists;
       }

       public static void CloseUnclaimedPropertyWindow()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           TitlebarActions.ClickClose((WinWindow)unclaimedProperty);
       }

       private static UITestControl GetTabs()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           var children = unclaimedProperty.GetChildren();
           return children[3];
       }

       public static void SelectPrintLetterTab()
       {
           var tabs = GetTabs();
           var tabList = tabs.Container.SearchFor<WinWindow>(new { ControlName = "ultraTabControl3" });
           var selectedTab = tabList.Container.SearchFor<WinTabPage>(new { Name = "InternalProcessPrintmethod" });
           MouseActions.Click(selectedTab);

       }

       public static void ClickOnGoBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty,PrintMethodConstants.GoBtn);
       }

       public static void ClickOnSelectAllBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty, PrintMethodConstants.SelectAllBtn);
       }

       public static void ClickOnSelectOneBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty, PrintMethodConstants.SelectOneBtn);
       }

       public static void ClickOnSaveAsExcelBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty, PrintMethodConstants.SaveasExcelBtn);
       }

       public static void ClickOnPrintLetterBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty, PrintMethodConstants.PrintLettersBtn);
       }

       public static void ClickOnChangeStatusBtnPrintLetter()
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();
           MouseActions.ClickButton(unclaimedProperty, PrintMethodConstants.ChangeStatusBtn);
       }

       public static void EnterDataInPrintLetterTab(DataRow data)
       {
           var unclaimedProperty = GetUnclaimedPropertyWindowProperties();

           var radioBtn = Actions.GetWindowChild(unclaimedProperty, PrintMethodConstants.CheckType);
           if (!string.IsNullOrEmpty(data.ItemArray[3].ToString()))
           {
               var radio = radioBtn.Container.SearchFor<WinRadioButton>(new { Name = data.ItemArray[3].ToString() });
               Actions.SelectRadioButton(radio);
           }

           var duration = Actions.GetWindowChild(unclaimedProperty, PrintMethodConstants.Duration);
           if(!string.IsNullOrEmpty(data.ItemArray[4].ToString()))
               DropDownActions.SelectDropdownByText(duration, data.ItemArray[4].ToString());

           MouseActions.ClickButton(unclaimedProperty,PrintMethodConstants.GoBtn);

           var totalAmt = Actions.GetWindowChild(unclaimedProperty, PrintMethodConstants.TotalAmt);
           if (totalAmt.Enabled)
           {
               if (!string.IsNullOrEmpty(data.ItemArray[5].ToString()))
                   Actions.SetText(totalAmt, data.ItemArray[5].ToString());
           }

           var minAmt = Actions.GetWindowChild(unclaimedProperty, PrintMethodConstants.MinimumAmt);
           if (!string.IsNullOrEmpty(data.ItemArray[6].ToString()))
               Actions.SetText(minAmt, data.ItemArray[6].ToString());

           var replyDt = Actions.GetWindowChild(unclaimedProperty, PrintMethodConstants.ReplyByDt);
           if (!string.IsNullOrEmpty(data.ItemArray[7].ToString()))
               DropDownActions.SelectDropdownByText(replyDt, data.ItemArray[7].ToString());
       }

       public static bool ClosePrintWindow()
       {
           var printWindow = GetprintWindowProperties();
           if (printWindow.Exists)
           {
               printWindow.SetFocus();
               SendKeys.SendWait("{ESC}");
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

       #region Controls

       private class PrintMethodConstants
       {
           public const string CheckType = "_optCheckType";
           public const string Duration = "ddlDuration";
           public const string GoBtn = "btnGo";
           public const string SelectAllBtn = "btnSelectAll";
           public const string SelectOneBtn = "btnSelectNone";
           public const string TotalAmt = "txtTotalAmount";
           public const string MinimumAmt = "txtMinimumAmount";
           public const string ReplyByDt = "dtReplybyDate";
           public const string SaveasExcelBtn = "btnSaveAsExcel";
           public const string PrintLettersBtn = "btnPrintLetters";
           public const string ChangeStatusBtn = "btnChangeStatus";
           public const string PrintLettersGrd = "grdPrintLetters";
       }

       #endregion
   }
}
