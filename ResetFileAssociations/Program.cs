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
            SetSystemDefaultBrowserWithGUI("Microsoft Edge");
        }

        public static bool SetSystemDefaultBrowserWithGUI(string aBrowserName)
        {
            Process.Start("ms-settings:about");

            var window = WaitForWindow("Settings");
            if (window == null)
                return false; // did not find in 10 seconds

            // find all the invokable stuff here
            var cond = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "Default apps"),
                new PropertyCondition(AutomationElement.IsSelectionItemPatternAvailableProperty, true));

            var defaultAppsButton = window.FindFirst(TreeScope.Descendants, cond);
            ((SelectionItemPattern)defaultAppsButton?.GetCurrentPattern(SelectionItemPattern.Pattern))?.Select();

            if (!WaitForLabel(window, "Choose default apps"))
                return false; // did not go as expected

            // find the reset button
            var resetBtnCond = new PropertyCondition(AutomationElement.NameProperty, "Reset");
            var resetBtn = window.FindFirst(TreeScope.Descendants, resetBtnCond)?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

            if (resetBtn == null)
                return false;

            resetBtn.Invoke();

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

        public static bool WaitForLabel(AutomationElement window, string title)
        {
            AutomationElement label = null;
            var condition = new PropertyCondition(AutomationElement.NameProperty, "Choose default apps");
            for (int attempts = 0; label == null && attempts < 100; Thread.Sleep(100), attempts++)
            {
                label = window.FindFirst(TreeScope.Descendants, condition);
            }

            return label != null;
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
    }
}
