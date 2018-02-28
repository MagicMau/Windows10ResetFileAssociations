using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace ResetFileAssociations
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Thread.Sleep(5000);
            if (WaitForUnlock())
            {
                SetSystemDefaultBrowserWithGUI();
            }
        }

        public static bool SetSystemDefaultBrowserWithGUI()
        {
            Process.Start("ms-settings:about");

            var window = WaitForWindow("Settings");
            if (window == null)
                return false; // did not find in 10 seconds

            // find all the invokable stuff here
            var cond = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "Default apps"),
                new PropertyCondition(AutomationElement.IsSelectionItemPatternAvailableProperty, true));

            var defaultAppsButton = WaitForCondition(window, cond);
            if (defaultAppsButton == null)
                return false;

            // make sure it's visible first
            bool isOffscreen = true;
            for (int attempts = 0; attempts < 100 && isOffscreen; attempts++)
            {
                Thread.Sleep(100);
                isOffscreen = (bool)defaultAppsButton.GetCurrentPropertyValue(AutomationElementIdentifiers.IsOffscreenProperty);
            }

            if (isOffscreen)
                return false;

            // now click it
            ((SelectionItemPattern)defaultAppsButton?.GetCurrentPattern(SelectionItemPattern.Pattern))?.Select();

            if (WaitForLabel(window, "Choose default apps") == null)
                return false; // did not go as expected

            var browserCond = new PropertyCondition(AutomationElement.NameProperty, "Web browser");
            var browserObj = WaitForCondition(window, browserCond);
            if (browserObj == null)
                return false;

            var next = TreeWalker.ControlViewWalker.GetNextSibling(browserObj);

            // the first child of next should be a text reading the name of the current browser
            var nextLabel = (TextPattern)TreeWalker.ControlViewWalker.GetFirstChild(next).GetCurrentPattern(TextPattern.Pattern);
            var curBrowser = nextLabel.DocumentRange.GetText(100);

            if (curBrowser == "Microsoft Edge")
            {
                ((WindowPattern)window.GetCurrentPattern(WindowPattern.Pattern)).Close();
                return true; // no need to change
            }

            // scroll down first
            var scrollWnd = TreeWalker.ControlViewWalker.GetParent(next);
            var scroll = (ScrollPattern)scrollWnd.GetCurrentPattern(ScrollPattern.Pattern);
            scroll.Scroll(ScrollAmount.NoAmount, ScrollAmount.LargeIncrement);

            // click the Web browser selection button
            var btn = ((InvokePattern)next.GetCurrentPattern(InvokePattern.Pattern));
            btn.Invoke();

            // this will open up the choices window - now how to find it...
            var edgeText = WaitForLabel(window, "Google Chrome");

            if (edgeText != null)
            {
                if (!edgeText.TryGetCurrentPattern(InvokePattern.Pattern, out object edgeBtnObj))
                {
                    var edgeTextParent = TreeWalker.ControlViewWalker.GetParent(edgeText);
                    if (!edgeTextParent.TryGetCurrentPattern(InvokePattern.Pattern, out edgeBtnObj))
                        return false; // oops
                }

                var edgeBtn = edgeBtnObj as InvokePattern;
                edgeBtn?.Invoke();
            }

            // close settings
            ((WindowPattern)window.GetCurrentPattern(WindowPattern.Pattern)).Close();

            return true;
        }


        public static AutomationElement WaitForWindow(string title)
        {
            AutomationElement mainWindow = null;
            for (int attempts = 0; mainWindow == null && attempts < 100; Thread.Sleep(100), attempts++)
            {
                mainWindow = GetSpecificWindow(title);
            }

            return mainWindow;
        }

        public static AutomationElement WaitForLabel(AutomationElement window, string title)
        {
            AutomationElement label = null;
            var condition = new PropertyCondition(AutomationElement.NameProperty, title);
            for (int attempts = 0; label == null && attempts < 100; Thread.Sleep(100), attempts++)
            {
                label = window.FindFirst(TreeScope.Descendants, condition);
            }

            return label;
        }

        public static AutomationElement WaitForCondition(AutomationElement window, Condition condition)
        {
            AutomationElement elt = null;
            for (int attempts = 0; elt == null && attempts < 100; Thread.Sleep(100), attempts++)
            {
                elt = window.FindFirst(TreeScope.Descendants, condition);
            }

            return elt;
        }


        public static AutomationElement GetSpecificWindow(string aWinTitle)
        {
            AutomationElement mainWindow = null;
            AutomationElementCollection winCollection = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);

            foreach (AutomationElement ele in winCollection)
            {
                if (ele.Current.Name.ToLower() == aWinTitle.ToLower())
                {
                    mainWindow = ele;
                    break;
                }
            }
            return mainWindow;
        }

        public static AutomationElement GetSpecificAutomationItem(string aWinTitle, string itemName)
        {
            AutomationElement window = GetSpecificWindow(aWinTitle);
            Condition condition = new PropertyCondition(AutomationElement.NameProperty, itemName);
            return window.FindFirst(TreeScope.Element | TreeScope.Descendants, condition);
        }

        private static bool WaitForUnlock()
        {
            SessionInfo.LockState state;
            int attempts = 0;

            // try for 3 hours
            while ((state = SessionInfo.GetSessionLockState(1)) == SessionInfo.LockState.Locked && attempts++ < 1080)
            {
                Trace.WriteLine($"Session is locked, attempt {attempts}.");
                // We are locked, so let's wait a bit
                Thread.Sleep(10000);
            }

            return attempts < 1080;
        }
    }
}
