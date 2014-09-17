//using EllisWinAppTest.Helpers;
//using Microsoft.VisualStudio.TestTools.UITest.Extension;
//using Microsoft.VisualStudio.TestTools.UITesting;
//using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

//namespace EllisWinAppTest.Elements
//{
//    public partial class WorkerProfile
//    {
//        //public void workersearch()
//        //{
//        //    #region Variable Declarations
//        //    WinCell uIWorkerCell = UISearchResultsWindow.UI_grdSearchResultWindow.UI_grdSearchResultTable.UIWorkerSearchResultDoRow.UIWorkerCell;
//        //    WinText uIWorkerText = UIWorkerProfileAALTONEWindow.UIWorkerWindow.UIWorkerText;
//        //    //WinText uIItem5141985Text = UIWorkerProfileAALTONEWindow.UIItem5141985Window.UIItem5141985Text;
//        //    //WinText uIXXXXX5828Text = UIWorkerProfileAALTONEWindow.UIXXXXX5828Window.UIXXXXX5828Text;
//        //    WinButton uICloseButton = UIWorkerProfileAALTONEWindow.UICloseWindow.UICloseButton;
//        //    WinButton uICloseButton1 = UISearchResultsWindow.UISearchResultsTitleBar.UICloseButton;
//        //    #endregion

//        //    // Double-Click 'AALTONEN, MATTHEW' cell
//        //    Mouse.DoubleClick(uIWorkerCell, new Point(13, 13));

//        //    // Click 'AALTONEN, MATTHEW' label
//        //    Mouse.Click(uIWorkerText, new Point(50, 3));

//        //    // Click '5/14/1985' label
//        //    //Mouse.Click(uIItem5141985Text, new Point(13, 11));

//        //    // Click 'XXX-XX-5828' label
//        //    //Mouse.Click(uIXXXXX5828Text, new Point(35, 8));

//        //    // Click 'Close' button
//        //    Mouse.Click(uICloseButton, new Point(33, 8));

//        //    // Click 'Close' button
//        //    Mouse.Click(uICloseButton1, new Point(13, 12));
//        //}

//        public static void SelectWorkerFromResultsWindow()
//        {
//            if (UISearchResultsWindow.UI_grdSearchResultWindow.UI_grdSearchResultTable.Exists)
//            {
//                var workerCell =
//                    UISearchResultsWindow.UI_grdSearchResultWindow.UI_grdSearchResultTable.UIWorkerSearchResultDoRow
//                        .UIWorkerCell;
//                var clickPoints = Factory.GetMouseCoOrdinates(workerCell);

//                //Mouse.DoubleClick(workerCell, new Point(clickPoints.X, clickPoints.Y));
//                Mouse.DoubleClick(clickPoints);
//            }
//        }

//        public static string GetWorkerProfileText()
//        {
//            var text = UIWorkerWindow.UIWorkerText;
//            return text.DisplayText;
//        }

//        public static void CloseWorkerProfile()
//        {
//            var control = UIWorkerProfileWindow.UICloseWindow.UICloseButton;

//            var clickPoints = Factory.GetMouseCoOrdinates(control);
//            Mouse.Click(clickPoints);
//        }

//        public static void CloseResultsWindow()
//        {
//            var control = UISearchResultsWindow.UISearchResultsTitleBar.UICloseButton;

//            var clickPoints = Factory.GetMouseCoOrdinates(control);
//            Mouse.Click(clickPoints);
//        }

//        public static UISearchResultsWindow UISearchResultsWindow
//        {
//            get
//            {
//                if ((mUISearchResultsWindow == null))
//                {
//                    mUISearchResultsWindow = new UISearchResultsWindow();
//                }
//                return mUISearchResultsWindow;
//            }
//        }

//        #region Fields

//        //    private workersearchParams mworkersearchParams;

//        //    private UIEllisWindow mUIEllisWindow;

//        //    private UIItemWindow mUIItemWindow;

//        //    private UIItemWindow1 mUIItemWindow1;

//        private static UISearchResultsWindow mUISearchResultsWindow;

//        private UIWorkerProfileWindow mUIWorkerProfileWindow;

//        #endregion
//    }

//    public class UISearchResultsWindow : WinWindow
//    {
//        public UISearchResultsWindow()
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.Name] = "Search Results";
//            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
//                PropertyExpressionOperator.Contains));
//            WindowTitles.Add("Search Results");

//            #endregion
//        }

//        #region Properties

//        public UI_grdSearchResultWindow UI_grdSearchResultWindow
//        {
//            get
//            {
//                if ((mUI_grdSearchResultWindow == null))
//                {
//                    mUI_grdSearchResultWindow = new UI_grdSearchResultWindow(this);
//                }
//                return mUI_grdSearchResultWindow;
//            }
//        }

//        public UISearchResultsTitleBar UISearchResultsTitleBar
//        {
//            get
//            {
//                if ((mUISearchResultsTitleBar == null))
//                {
//                    mUISearchResultsTitleBar = new UISearchResultsTitleBar(this);
//                }
//                return mUISearchResultsTitleBar;
//            }
//        }

//        #endregion

//        #region Fields

//        private UI_grdSearchResultWindow mUI_grdSearchResultWindow;

//        private UISearchResultsTitleBar mUISearchResultsTitleBar;

//        #endregion
//    }

//    public class UI_grdSearchResultWindow : WinWindow
//    {
//        public UI_grdSearchResultWindow(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "_grdSearchResult";
//            WindowTitles.Add("Search Results");

//            #endregion
//        }

//        #region Properties

//        public UI_grdSearchResultTable UI_grdSearchResultTable
//        {
//            get
//            {
//                if ((mUI_grdSearchResultTable == null))
//                {
//                    mUI_grdSearchResultTable = new UI_grdSearchResultTable(this);
//                }
//                return mUI_grdSearchResultTable;
//            }
//        }

//        #endregion

//        #region Fields

//        private UI_grdSearchResultTable mUI_grdSearchResultTable;

//        #endregion
//    }

//    public class UI_grdSearchResultTable : WinTable
//    {
//        public UI_grdSearchResultTable(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            WindowTitles.Add("Search Results");

//            #endregion
//        }

//        #region Properties

//        public UIWorkerSearchResultDoRow UIWorkerSearchResultDoRow
//        {
//            get
//            {
//                if ((mUIWorkerSearchResultDoRow == null))
//                {
//                    mUIWorkerSearchResultDoRow = new UIWorkerSearchResultDoRow(this);
//                }
//                return mUIWorkerSearchResultDoRow;
//            }
//        }

//        #endregion

//        #region Fields

//        private UIWorkerSearchResultDoRow mUIWorkerSearchResultDoRow;

//        #endregion
//    }

//    public class UIWorkerSearchResultDoRow : WinRow
//    {
//        public UIWorkerSearchResultDoRow(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.Name] = "WorkerSearchResultDomain row 1";
//            SearchConfigurations.Add(SearchConfiguration.AlwaysSearch);
//            WindowTitles.Add("Search Results");

//            #endregion
//        }

//        #region Properties

//        public WinCell UIWorkerCell
//        {
//            get
//            {
//                if ((mUIWorkerCell == null))
//                {
//                    mUIWorkerCell = new WinCell(this);

//                    #region Search Criteria

//                    mUIWorkerCell.SearchProperties[WinCell.PropertyNames.Value] = Globals.WorkerName;
//                    mUIWorkerCell.WindowTitles.Add("Search Results");

//                    #endregion
//                }
//                return mUIWorkerCell;
//            }
//        }

//        #endregion

//        #region Fields

//        private WinCell mUIWorkerCell;

//        #endregion
//    }

//    public class UISearchResultsTitleBar : WinTitleBar
//    {
//        public UISearchResultsTitleBar(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            WindowTitles.Add("Search Results");

//            #endregion
//        }

//        #region Properties

//        public WinButton UICloseButton
//        {
//            get
//            {
//                if ((mUICloseButton == null))
//                {
//                    mUICloseButton = new WinButton(this);

//                    #region Search Criteria

//                    mUICloseButton.SearchProperties[WinButton.PropertyNames.Name] = "Close";
//                    mUICloseButton.WindowTitles.Add("Search Results");

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

//    public class UIWorkerProfileWindow : WinWindow
//    {
//        public UIWorkerProfileWindow()
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.Name] = "Worker Profile-" + Globals.WorkerName;
//            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window",
//                PropertyExpressionOperator.Contains));
//            WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//            #endregion
//        }

//        #region Properties

//        public static UIWorkerWindow UIWorkerWindow
//        {
//            get
//            {
//                if ((mUIWorkerWindow == null))
//                {
//                    UIWorkerProfileWindow all = new UIWorkerProfileWindow();
//                    mUIWorkerWindow = new UIWorkerWindow(all);
//                }
//                return mUIWorkerWindow;
//            }
//        }

//        //public UIItem5141985Window UIItem5141985Window
//        //{
//        //    get
//        //    {
//        //        if ((mUIItem5141985Window == null))
//        //        {
//        //            mUIItem5141985Window = new UIItem5141985Window(this);
//        //        }
//        //        return mUIItem5141985Window;
//        //    }
//        //}

//        //public UIXXXXX5828Window UIXXXXX5828Window
//        //{
//        //    get
//        //    {
//        //        if ((mUIXXXXX5828Window == null))
//        //        {
//        //            mUIXXXXX5828Window = new UIXXXXX5828Window(this);
//        //        }
//        //        return mUIXXXXX5828Window;
//        //    }
//        //}

//        public static UICloseWindow UICloseWindow
//        {
//            get
//            {
//                if ((mUICloseWindow == null))
//                {
//                    //mUICloseWindow = new UICloseWindow(this);
//                    mUICloseWindow = new UICloseWindow();
//                }
//                return mUICloseWindow;
//            }
//        }

//        #endregion

//        #region Fields

//        private static UIWorkerWindow mUIWorkerWindow;

//        //private UIItem5141985Window mUIItem5141985Window;

//        //private UIXXXXX5828Window mUIXXXXX5828Window;

//        private static UICloseWindow mUICloseWindow;

//        #endregion
//    }

//    public class UIWorkerWindow : WinWindow
//    {
//        public UIWorkerWindow(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "lblName";
//            WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//            #endregion
//        }

//        public UIWorkerWindow()
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "lblName";
//            WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//            #endregion
//        }

//        #region Properties

//        public static WinText UIWorkerText
//        {
//            get
//            {
//                if ((mUIWorkerText == null))
//                {
//                    mUIWorkerText = new WinText();

//                    #region Search Criteria

//                    mUIWorkerText.SearchProperties[WinText.PropertyNames.Name] = Globals.WorkerName;
//                    mUIWorkerText.WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//                    #endregion
//                }
//                return mUIWorkerText;
//            }
//        }

//        #endregion

//        #region Fields

//        private static WinText mUIWorkerText;

//        #endregion
//    }

//    //public class UIItem5141985Window : WinWindow
//    //{

//    //    public UIItem5141985Window(UITestControl searchLimitContainer) :
//    //        base(searchLimitContainer)
//    //    {
//    //        #region Search Criteria
//    //        SearchProperties[PropertyNames.ControlName] = "lblDateOfBirthText";
//    //        WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");
//    //        #endregion
//    //    }

//    //    #region Properties
//    //    public WinText UIItem5141985Text
//    //    {
//    //        get
//    //        {
//    //            if ((mUIItem5141985Text == null))
//    //            {
//    //                mUIItem5141985Text = new WinText(this);
//    //                #region Search Criteria
//    //                mUIItem5141985Text.SearchProperties[WinText.PropertyNames.Name] = "5/14/1985";
//    //                mUIItem5141985Text.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");
//    //                #endregion
//    //            }
//    //            return mUIItem5141985Text;
//    //        }
//    //    }
//    //    #endregion

//    //    #region Fields
//    //    private WinText mUIItem5141985Text;
//    //    #endregion
//    //}

//    //public class UIXXXXX5828Window : WinWindow
//    //{

//    //    public UIXXXXX5828Window(UITestControl searchLimitContainer) :
//    //        base(searchLimitContainer)
//    //    {
//    //        #region Search Criteria
//    //        SearchProperties[PropertyNames.ControlName] = "lblSSNnoText";
//    //        WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");
//    //        #endregion
//    //    }

//    //    #region Properties
//    //    public WinText UIXXXXX5828Text
//    //    {
//    //        get
//    //        {
//    //            if ((mUIXXXXX5828Text == null))
//    //            {
//    //                mUIXXXXX5828Text = new WinText(this);
//    //                #region Search Criteria
//    //                mUIXXXXX5828Text.SearchProperties[WinText.PropertyNames.Name] = "XXX-XX-5828";
//    //                mUIXXXXX5828Text.WindowTitles.Add("Worker Profile-AALTONEN, MATTHEW");
//    //                #endregion
//    //            }
//    //            return mUIXXXXX5828Text;
//    //        }
//    //    }
//    //    #endregion

//    //    #region Fields
//    //    private WinText mUIXXXXX5828Text;
//    //    #endregion
//    //}

//    public class UICloseWindow : WinWindow
//    {
//        public UICloseWindow(UITestControl searchLimitContainer) :
//            base(searchLimitContainer)
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "btnClose";
//            WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//            #endregion
//        }

//        public UICloseWindow()
//        {
//            #region Search Criteria

//            SearchProperties[PropertyNames.ControlName] = "btnClose";
//            WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

//            #endregion
//        }

//        #region Properties

//        public WinButton UICloseButton
//        {
//            get
//            {
//                if ((mUICloseButton == null))
//                {
//                    mUICloseButton = new WinButton(this);

//                    #region Search Criteria

//                    mUICloseButton.SearchProperties[WinButton.PropertyNames.Name] = "Close";
//                    mUICloseButton.WindowTitles.Add("Worker Profile-" + Globals.WorkerName);

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