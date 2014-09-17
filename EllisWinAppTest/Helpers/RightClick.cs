using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Input;
using Ellis.WinApp.Testing.Framework;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Helpers
{
    public class RightClick
    {
        public static bool Reschedule()
        {
            try
            {
                #region Variable Declarations

                var uIALLFLIGHTCORPORATIONCell =
                    UIEllisWindow.UI_gridOverdueCustomerWindow.UI_gridOverdueCustomerTable.UIOverdueCustomerSummaRow
                        .UIALLFLIGHTCORPORATIONCell;
                var uIRescheduleMenuItem = UIItemWindow.UIDropDownMenu.UIRescheduleMenuItem;
                //var uICancelCallBackMenuItem = UIItemWindow.UIDropDownMenu.UICancelCallBackMenuItem;

                #endregion

                uIALLFLIGHTCORPORATIONCell.WaitForControlReady();
                uIALLFLIGHTCORPORATIONCell.SetFocus();
                var clickPoints =
                    new Point(
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Width / 2 +
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.X,
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Height / 2 +
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Y);
                Mouse.Click(uIALLFLIGHTCORPORATIONCell);
                Mouse.Click(uIALLFLIGHTCORPORATIONCell, MouseButtons.Right, ModifierKeys.None,
                    new Point(clickPoints.X, clickPoints.Y));

                // Click 'Reschedule' menu item
                Mouse.Click(uIRescheduleMenuItem);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool CallBack()
        {
            try
            {
                #region Variable Declarations
                var uIALLFLIGHTCORPORATIONCell =
               UIEllisWindow.UI_gridOverdueCustomerWindow.UI_gridOverdueCustomerTable.UIOverdueCustomerSummaRow
                   .UIALLFLIGHTCORPORATIONCell;
                //var uIRescheduleMenuItem = UIItemWindow.UIDropDownMenu.UIRescheduleMenuItem;
                var uICancelCallBackMenuItem = UIItemWindow.UIDropDownMenu.UICancelCallBackMenuItem;

            #endregion

                uIALLFLIGHTCORPORATIONCell.WaitForControlReady();
                uIALLFLIGHTCORPORATIONCell.SetFocus();
                var clickPoints =
                    new Point(
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Width / 2 +
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.X,
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Height / 2 +
                        uIALLFLIGHTCORPORATIONCell.BoundingRectangle.Y);
                Mouse.Click(uIALLFLIGHTCORPORATIONCell);
                Mouse.Click(uIALLFLIGHTCORPORATIONCell, MouseButtons.Right, ModifierKeys.None,
                    new Point(clickPoints.X, clickPoints.Y));

                // Click 'Cancel Call Back' menu item
                Mouse.Click(uICancelCallBackMenuItem);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
           
        }

        #region Properties

        public static UIEllisWindow UIEllisWindow
        {
            get
            {
                if ((mUIEllisWindow == null))
                {
                    mUIEllisWindow = new UIEllisWindow();
                }
                return mUIEllisWindow;
            }
        }

        public static DDItemWindow UIItemWindow
        {
            get
            {
                if ((mUIItemWindow == null))
                {
                    mUIItemWindow = new DDItemWindow();
                }
                return mUIItemWindow;
            }
        }

        #endregion

        #region Fields

        private static UIEllisWindow mUIEllisWindow;

        private static DDItemWindow mUIItemWindow;

        #endregion
    }

    public class UIEllisWindow : WinWindow
    {
        public UIEllisWindow()
        {
            #region Search Criteria

            SearchProperties[PropertyNames.Name] = "Ellis";
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
                PropertyExpressionOperator.Contains));
            WindowTitles.Add("Ellis");

            #endregion
        }

        #region Properties

        public UI_gridOverdueCustomerWindow UI_gridOverdueCustomerWindow
        {
            get
            {
                if ((mUI_gridOverdueCustomerWindow == null))
                {
                    mUI_gridOverdueCustomerWindow = new UI_gridOverdueCustomerWindow(this);
                }
                return mUI_gridOverdueCustomerWindow;
            }
        }

        #endregion

        #region Fields

        private UI_gridOverdueCustomerWindow mUI_gridOverdueCustomerWindow;

        #endregion
    }

    public class UI_gridOverdueCustomerWindow : WinWindow
    {
        public UI_gridOverdueCustomerWindow(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            SearchProperties[PropertyNames.ControlName] = "_gridOverdueCustomers";
            WindowTitles.Add("Ellis");

            #endregion
        }

        #region Properties

        public UI_gridOverdueCustomerTable UI_gridOverdueCustomerTable
        {
            get
            {
                if ((mUI_gridOverdueCustomerTable == null))
                {
                    mUI_gridOverdueCustomerTable = new UI_gridOverdueCustomerTable(this);
                }
                return mUI_gridOverdueCustomerTable;
            }
        }

        #endregion

        #region Fields

        private UI_gridOverdueCustomerTable mUI_gridOverdueCustomerTable;

        #endregion
    }

    public class UI_gridOverdueCustomerTable : WinTable
    {
        public UI_gridOverdueCustomerTable(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            WindowTitles.Add("Ellis");

            #endregion
        }

        #region Properties

        public UIOverdueCustomerSummaRow UIOverdueCustomerSummaRow
        {
            get
            {
                if ((mUIOverdueCustomerSummaRow == null))
                {
                    mUIOverdueCustomerSummaRow = new UIOverdueCustomerSummaRow(this);
                }
                return mUIOverdueCustomerSummaRow;
            }
        }

        #endregion

        #region Fields

        private UIOverdueCustomerSummaRow mUIOverdueCustomerSummaRow;

        #endregion
    }

    public class UIOverdueCustomerSummaRow : WinRow
    {
        public UIOverdueCustomerSummaRow(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            var searchCriteria = "OverdueCustomerSummaryDomain row " + Generator.GenerateRandomNumber(35);
            SearchProperties[PropertyNames.Name] = searchCriteria;
            SearchConfigurations.Add(SearchConfiguration.AlwaysSearch);
            WindowTitles.Add("Ellis");

            #endregion
        }

        #region Properties

        public WinCell UIALLFLIGHTCORPORATIONCell
        {
            get
            {
                if ((mUIALLFLIGHTCORPORATIONCell == null))
                {
                    mUIALLFLIGHTCORPORATIONCell = new WinCell(this);

                    #region Search Criteria

                    mUIALLFLIGHTCORPORATIONCell.SearchProperties[WinCell.PropertyNames.Name] = Globals.Temp;
                    mUIALLFLIGHTCORPORATIONCell.WindowTitles.Add("Ellis");
                    //Globals.Temp = mUIALLFLIGHTCORPORATIONCell.Value;

                    #endregion
                }
                return mUIALLFLIGHTCORPORATIONCell;
            }
        }

        #endregion

        #region Fields

        private WinCell mUIALLFLIGHTCORPORATIONCell;

        #endregion
    }

    public class DDItemWindow : WinWindow
    {
        public DDItemWindow()
        {
            #region Search Criteria

            SearchProperties[PropertyNames.AccessibleName] = "DropDown";
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
                PropertyExpressionOperator.Contains));

            #endregion
        }

        #region Properties

        public UIDropDownMenu UIDropDownMenu
        {
            get
            {
                if ((mUIDropDownMenu == null))
                {
                    mUIDropDownMenu = new UIDropDownMenu(this);
                }
                return mUIDropDownMenu;
            }
        }

        #endregion

        #region Fields

        private UIDropDownMenu mUIDropDownMenu;

        #endregion
    }

    public class UIDropDownMenu : WinMenu
    {
        public UIDropDownMenu(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            SearchProperties[PropertyNames.Name] = "DropDown";

            #endregion
        }

        #region Properties

        public WinMenuItem UIRescheduleMenuItem
        {
            get
            {
                if ((mUIRescheduleMenuItem == null))
                {
                    mUIRescheduleMenuItem = new WinMenuItem(this);

                    #region Search Criteria

                    mUIRescheduleMenuItem.SearchProperties[WinMenuItem.PropertyNames.Name] = "Reschedule";

                    #endregion
                }
                return mUIRescheduleMenuItem;
            }
        }

        public WinMenuItem UICancelCallBackMenuItem
        {
            get
            {
                if ((mUICancelCallBackMenuItem == null))
                {
                    mUICancelCallBackMenuItem = new WinMenuItem(this);

                    #region Search Criteria

                    mUICancelCallBackMenuItem.SearchProperties[WinMenuItem.PropertyNames.Name] = "Cancel Call Back";

                    #endregion
                }
                return mUICancelCallBackMenuItem;
            }
        }

        #endregion

        #region Fields

        private WinMenuItem mUIRescheduleMenuItem;

        private WinMenuItem mUICancelCallBackMenuItem;

        #endregion
    }
}