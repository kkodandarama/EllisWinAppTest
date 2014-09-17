//using System.Drawing;
//using EllisWinAppTest.Helpers;
//using Microsoft.VisualStudio.TestTools.UITesting;
//using Microsoft.VisualStudio.TestTools.UITesting.WinControls;


//namespace EllisWinAppTest.Elements
//{
//    public partial class CustomerProfile
//    {
//        public static string GetCustomerProfileText()
//        {
//            var text = CustomerProfileWindow.CpWindow.CpWindowText;
//            return text.DisplayText;
//        }

//        public static void CloseCustomerProfile()
//        {
//            var control = CustomerProfileWindow.CustomerProfileTitleBar.CpCloseButton;
//            //control.SetFocus();
//            var clickPoints = new Point(control.BoundingRectangle.Width/2 + control.BoundingRectangle.X,
//                control.BoundingRectangle.Height/2 + control.BoundingRectangle.Y);

//            Mouse.Click(clickPoints);
//        }

//        #region Properties

//        public static UICustomerCustomeWindow CustomerProfileWindow
//        {
//            get
//            {
//                if ((mUICustomerCustomeWindow == null))
//                {
//                    mUICustomerCustomeWindow = new UICustomerCustomeWindow();
//                }
//                return mUICustomerCustomeWindow;
//            }
//        }

//        #endregion

//        private static UICustomerCustomeWindow mUICustomerCustomeWindow;
//    }

//    public class UICustomerCustomeWindow : WinWindow
//    {
//        public UICustomerCustomeWindow()
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.Name] = Globals.CustomerName + " - Customer Profile";
//            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
//                PropertyExpressionOperator.Contains));
//            WindowTitles.Add(Globals.CustomerName + " - Customer Profile");

//            #endregion
//        }

//        #region Properties

//        public UICustomerWindow CpWindow
//        {
//            get
//            {
//                if ((mUICustomerWindow == null))
//                {
//                    mUICustomerWindow = new UICustomerWindow(this);
//                }
//                return mUICustomerWindow;
//            }
//        }

//        public UICustomerCustomeTitleBar CustomerProfileTitleBar
//        {
//            get
//            {
//                if ((mUICustomerCustomeTitleBar == null))
//                {
//                    mUICustomerCustomeTitleBar = new UICustomerCustomeTitleBar(this);
//                }
//                return mUICustomerCustomeTitleBar;
//            }
//        }

//        #endregion

//        #region Fields

//        private UICustomerWindow mUICustomerWindow;

//        private UICustomerCustomeTitleBar mUICustomerCustomeTitleBar;

//        #endregion
//    }

//    public class UICustomerWindow : WinWindow
//    {
//        public UICustomerWindow(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "CaseTitleLabel";
//            WindowTitles.Add(Globals.CustomerName + " - Customer Profile");

//            #endregion
//        }

//        #region Properties

//        public WinText CpWindowText
//        {
//            get
//            {
//                if ((mUICustomerText == null))
//                {
//                    mUICustomerText = new WinText(this);

//                    #region Search Criteria

//                    mUICustomerText.SearchProperties[WinText.PropertyNames.Name] = Globals.CustomerName;
//                    mUICustomerText.WindowTitles.Add(Globals.CustomerName + " - Customer Profile");

//                    #endregion
//                }
//                return mUICustomerText;
//            }
//        }

//        #endregion

//        #region Fields

//        private WinText mUICustomerText;

//        #endregion
//    }


//    public class UICustomerCustomeTitleBar : WinTitleBar
//    {
//        public UICustomerCustomeTitleBar(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            WindowTitles.Add(Globals.CustomerName + " - Customer Profile");

//            #endregion
//        }

//        #region Properties

//        public WinButton CpCloseButton
//        {
//            get
//            {
//                if ((mUICloseButton == null))
//                {
//                    mUICloseButton = new WinButton(this);

//                    #region Search Criteria

//                    mUICloseButton.SearchProperties[WinButton.PropertyNames.Name] = "Close";
//                    mUICloseButton.WindowTitles.Add(Globals.CustomerName + " - Customer Profile");

//                    #endregion
//                }
//                return mUICloseButton;
//            }
//        }

//        #endregion

//        #region Fields

//        private WinButton mUICloseButton;

//        #endregion
//    }
//}