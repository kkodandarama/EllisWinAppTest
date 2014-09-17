using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;

namespace EllisWinAppTest.Helpers
{
    public static class Extentions
    {
        /// <summary>
        /// This is a slow search, but if you can't find a control, it's a must have!  
        /// Send it your UITestControl root and any predicate you can think of.  
        /// It'll give you an IEnumerable of all matching UITestControls.
        /// </summary>
        /// <param name="predicate"></param>
        /// <returns>IEnumerable of all UITestControls matching your predicate.</returns>
        public static IEnumerable<UITestControl> BreadthFirstSearch(this UITestControl root, Func<UITestControl, bool> predicate)
        {
            var queue = new Queue<UITestControl>();
            queue.Enqueue(root);
            while (queue.Count > 0)
            {
                var child = queue.Dequeue();
                if (predicate(child))
                    yield return child;
                try
                {
                    foreach (var childControl in child.GetChildren())
                    {
                        queue.Enqueue(childControl);
                    }
                }
                catch
                {
                    Debug.WriteLine("A child threw during the BreadthFirstSearch.");
                }
            }

            yield break;
        }

        internal static bool IsClickablePointOnScreen(UITestControl uit)
        {
            if (uit == null) return false;

            try
            {
                return Screen.AllScreens.Any(
                    x => x.WorkingArea.Contains(
                        uit.GetClickablePoint()));
            }
            catch (NoClickablePointException ncp)
            {
                Debug.WriteLine(string.Format("The UITestControl doesn't have a clickable point. {0}", ncp));
            }
            catch (FailedToPerformActionOnBlockedControlException ena)
            {
                Debug.WriteLine(string.Format("The point is blocked by another window. This is a non-blocking error. \n{0}", ena));
            }

            return false;
        }

        internal static bool DoesControlIntersectScreen(UITestControl uit)
        {
            if (uit == null || uit.BoundingRectangle == null)
                return false;

            if (uit.BoundingRectangle.Right < Screen.AllScreens.Min(x => x.WorkingArea.Left) ||
                uit.BoundingRectangle.Left > Screen.AllScreens.Max(x => x.WorkingArea.Right) ||
                uit.BoundingRectangle.Bottom < Screen.AllScreens.Min(x => x.WorkingArea.Top) ||
                uit.BoundingRectangle.Top > Screen.AllScreens.Max(x => x.WorkingArea.Bottom))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// This is a much better test than UITestControl.Exists.  Many of the Ellis controls Exist, but are hidden off screen, are behind a modal dialog, or are generally not usable.
        /// This tests for usability in a much better way.
        /// </summary>
        /// <returns>True if the UITestControl intersects the user window and is a clickable item.</returns>
        public static bool IsControlUsable(this UITestControl uit)
        {
            try
            {
                return uit.Exists && DoesControlIntersectScreen(uit) && IsClickablePointOnScreen(uit);
            }
            catch (UITestControlNotFoundException notFound)
            {
                Debug.WriteLine(notFound);
                return false;
            }
            catch(UITestControlNotAvailableException notAvailable)
            {
                Debug.WriteLine(notAvailable);
                return false;
            }
        }
    }
}
