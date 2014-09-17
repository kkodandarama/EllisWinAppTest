using System.CodeDom.Compiler;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace EllisWinAppTest.Elements
{
    public class SelectRadioButton
    {
        public static void Selection(string option)
        {
            #region Variable Declarations

            switch (option)
            {
                case "Paycard":
                    {
                        var uIPaycardRadioButton =
                            UIWorkerProfileAALTONEWindow.UIPaycardWindow.UIItemGroup.UIPaycardRadioButton;
                        Mouse.Click(uIPaycardRadioButton);
                    }
                    break;
                case "Direct Deposit":
                    {
                        var uIDirectDepositRadioButton =
                            UIWorkerProfileAALTONEWindow.UIPaycardWindow.UIItemGroup.UIDirectDepositRadioButton;
                        Mouse.Click(uIDirectDepositRadioButton);
                    }
                    break;
                case "Check":
                    {
                        var uICheckRadioButton = UIWorkerProfileAALTONEWindow.UIPaycardWindow.UIItemGroup.UICheckRadioButton;
                        Mouse.Click(uICheckRadioButton);
                    }
                    break;
                case "Voucher":
                    {
                        var uIVoucherRadioButton =
                            UIWorkerProfileAALTONEWindow.UIPaycardWindow.UIItemGroup.UIVoucherRadioButton;
                        Mouse.Click(uIVoucherRadioButton);
                    }
                    break;
            }

            #endregion
        }

        #region Properties

        public static UIWorkerProfileAALTONEWindow UIWorkerProfileAALTONEWindow
        {
            get
            {
                if ((mUIWorkerProfileAALTONEWindow == null))
                {
                    mUIWorkerProfileAALTONEWindow = new UIWorkerProfileAALTONEWindow();
                }
                return mUIWorkerProfileAALTONEWindow;
            }
        }

        #endregion

        #region Fields

        private static UIWorkerProfileAALTONEWindow mUIWorkerProfileAALTONEWindow;

        #endregion
    }

    [GeneratedCode("Coded UITest Builder", "12.0.21005.1")]
    public class UIWorkerProfileAALTONEWindow : WinWindow
    {
        public UIWorkerProfileAALTONEWindow()
        {
            #region Search Criteria

            SearchProperties[PropertyNames.Name] = "Worker Profile-AALTONEN, MATTHEW";
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
                PropertyExpressionOperator.Contains));
            WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

            #endregion
        }

        #region Properties

        public UIPaycardWindow UIPaycardWindow
        {
            get
            {
                if ((mUIPaycardWindow == null))
                {
                    mUIPaycardWindow = new UIPaycardWindow(this);
                }
                return mUIPaycardWindow;
            }
        }

        #endregion

        #region Fields

        private UIPaycardWindow mUIPaycardWindow;

        #endregion
    }

    [GeneratedCode("Coded UITest Builder", "12.0.21005.1")]
    public class UIPaycardWindow : WinWindow
    {
        public UIPaycardWindow(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            SearchProperties[PropertyNames.ControlName] = "optPaymentMethod";
            WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

            #endregion
        }

        #region Properties

        public UIItemGroup UIItemGroup
        {
            get
            {
                if ((mUIItemGroup == null))
                {
                    mUIItemGroup = new UIItemGroup(this);
                }
                return mUIItemGroup;
            }
        }

        #endregion

        #region Fields

        private UIItemGroup mUIItemGroup;

        #endregion
    }

    [GeneratedCode("Coded UITest Builder", "12.0.21005.1")]
    public class UIItemGroup : WinGroup
    {
        public UIItemGroup(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria

            WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

            #endregion
        }

        #region Properties

        public WinRadioButton UIPaycardRadioButton
        {
            get
            {
                if ((mUIPaycardRadioButton == null))
                {
                    mUIPaycardRadioButton = new WinRadioButton(this);

                    #region Search Criteria

                    mUIPaycardRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = "Paycard";
                    mUIPaycardRadioButton.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

                    #endregion
                }
                return mUIPaycardRadioButton;
            }
        }

        public WinRadioButton UIDirectDepositRadioButton
        {
            get
            {
                if ((mUIDirectDepositRadioButton == null))
                {
                    mUIDirectDepositRadioButton = new WinRadioButton(this);

                    #region Search Criteria

                    mUIDirectDepositRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = "Direct Deposit";
                    mUIDirectDepositRadioButton.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

                    #endregion
                }
                return mUIDirectDepositRadioButton;
            }
        }

        public WinRadioButton UICheckRadioButton
        {
            get
            {
                if ((mUICheckRadioButton == null))
                {
                    mUICheckRadioButton = new WinRadioButton(this);

                    #region Search Criteria

                    mUICheckRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = "Check";
                    mUICheckRadioButton.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

                    #endregion
                }
                return mUICheckRadioButton;
            }
        }

        public WinRadioButton UIVoucherRadioButton
        {
            get
            {
                if ((mUIVoucherRadioButton == null))
                {
                    mUIVoucherRadioButton = new WinRadioButton(this);

                    #region Search Criteria

                    mUIVoucherRadioButton.SearchProperties[WinRadioButton.PropertyNames.Name] = "Voucher";
                    mUIVoucherRadioButton.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");

                    #endregion
                }
                return mUIVoucherRadioButton;
            }
        }

        #endregion

        #region Fields

        private WinRadioButton mUIPaycardRadioButton;

        private WinRadioButton mUIDirectDepositRadioButton;

        private WinRadioButton mUICheckRadioButton;

        private WinRadioButton mUIVoucherRadioButton;

        #endregion
    }
}