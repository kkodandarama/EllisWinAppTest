using System;
using System.Data;
using System.Windows.Forms;
using Ellis.WinApp.Testing.Framework;
using Ellis.WinApp.Testing.Framework.Actions;
using EllisWinAppTest.Helpers;
using EllisWinAppTest.Windows.EllisWindow;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;


namespace EllisWinAppTest.Windows.JobOrderWindow
{
    class SafetyWindow : AppContext
    {

        public static UITestControl JobOrderProfileWindowProperties()
        {
            var joborderWindow =
                App.Container.SearchFor<WinWindow>(new { ClassName = "WindowsForms10.Window.8.app.0.265601d" });
            return joborderWindow;
        }
        public static void SetSafetyDataOfJobOrder(DataRow data)
        {
            var windowInst = JobOrderProfileWindowProperties();
            if (windowInst.Exists)
            {
                if (data.ItemArray[82].ToString() == "Yes" || data.ItemArray[82].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[82].ToString(), "_optWrittenSafetyProgram");
                if (data.ItemArray[83].ToString() == "Yes" || data.ItemArray[83].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[83].ToString(), "_optCustomerSafetyProgram");
                if (data.ItemArray[84].ToString() == "Yes" || data.ItemArray[84].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[84].ToString(), "_optPPE");
                if (data.ItemArray[85].ToString() == "Yes" || data.ItemArray[85].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[85].ToString(), "_optSpecialisedTraining");
                if (data.ItemArray[86].ToString() == "Yes" || data.ItemArray[86].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[86].ToString(), "_optElevatedSurface");
                if (data.ItemArray[87].ToString() == "Yes" || data.ItemArray[87].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[87].ToString(), "_optHazardousMaterials");
                if (data.ItemArray[88].ToString() == "Yes" || data.ItemArray[88].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[88].ToString(), "_optTrenchWork");
                if (data.ItemArray[89].ToString() == "Yes" || data.ItemArray[89].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[89].ToString(), "_optConfinedSpace");
                if (data.ItemArray[90].ToString() == "Yes" || data.ItemArray[90].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[90].ToString(), "_optMovingMachinery");
                if (data.ItemArray[91].ToString() == "Yes" || data.ItemArray[91].ToString() == "No")
                    Factory.SelectRadioButton(windowInst, data.ItemArray[91].ToString(), "_optWorkersOnsite");

                var winControl = Actions.GetWindowChild(windowInst, "_dtpSafetyEvaluationDate");
                if (winControl.Enabled)
                {
                    winControl.SetFocus();
                    SendKeys.SendWait("+{END}");
                    SendKeys.SendWait("{DEL}");
                    SendKeys.SendWait("HOME");
                    SendKeys.SendWait(data.ItemArray[95].ToString());
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
                {
                    winControl = Actions.GetWindowChild(windowInst, "_dtpSiteEvaluated");
                    winControl.SetFocus();
                    SendKeys.SendWait("+{END}");
                    SendKeys.SendWait("{DEL}");
                    SendKeys.SendWait("HOME");
                    SendKeys.SendWait(data.ItemArray[95].ToString());
                }
                if (data.ItemArray[96].ToString() != "")
                {
                    var siteEvaluated =
                        windowInst.Container.SearchFor<WinRadioButton>(new { Name = "Satisfactory" });
                    Actions.SelectRadioButton(siteEvaluated);
                }

                MouseActions.ClickButton(windowInst, "btnAdd");
                MouseActions.ClickButton(windowInst, "btnSave");
                var alertWindow = windowInst.Container.SearchFor<WinWindow>(new { Name = "Alert" });
                MouseActions.ClickButton(alertWindow, "_OKButton");
            }
        }
    }
}
