namespace OutlookCustomPrintTest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    class BackStageTool
    {
        public static IntPtr GetActiveWindowHwnd()
        {
            dynamic w = Globals.ThisAddIn.Application.ActiveWindow();
            string caption = w.Caption;
            int top = w.Top;
            int left = w.Left;
            int width = w.Width;
            int height = w.Height;
            try
            {
                var list = Win32.EnumWindows();
                var founds = list.Where(_ =>
                {
                    var rect = Win32.GetWindowRect(_);
                    if (rect.Left == left
                        && rect.Top == top
                        && rect.Right - rect.Left == width)
                    {
                        if (Win32.GetWindowText(_) == caption)
                        {
                            return true;
                        }
                    }
                    return false;

                }).ToArray();
                return founds.FirstOrDefault();
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        class BackStageItem
        {
            public System.Windows.Automation.AutomationElement wnd;
            public System.Windows.Automation.AutomationElement BackStageView;
            public System.Windows.Automation.AutomationElement BackStage;
            public readonly List<System.Windows.Automation.AutomationElement> Columns = new List<System.Windows.Automation.AutomationElement>();

            public BackStageRect ToBackStageRect()
            {
                BackStageRect ret = new BackStageRect();
                ret.BackStageView = this.BackStageView.Current.BoundingRectangle;
                ret.BackStage = this.BackStage.Current.BoundingRectangle;
                foreach (var e in this.Columns)
                {
                    ret.Columns.Add(e.Current.BoundingRectangle);
                }
                return ret;
            }

            public static BackStageItem GetBackStageItem(IntPtr hwnd)
            {
                BackStageItem ret = new BackStageItem();
                try
                {
                    ret.wnd = System.Windows.Automation.AutomationElement.FromHandle(hwnd);
                    System.Windows.Automation.PropertyCondition conBackstageVie
                        = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.AutomationIdProperty, "BackstageView");
                    ret.BackStageView = ret.wnd.FindFirst(System.Windows.Automation.TreeScope.Subtree, conBackstageVie);

                    var conGroup = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Group);
                    ret.BackStage = ret.BackStageView.FindFirst(System.Windows.Automation.TreeScope.Subtree, conGroup);

                    ///2013と2016で違う？
                    var conPane = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Pane);
                    var conGroupOrPane = new System.Windows.Automation.OrCondition(new System.Windows.Automation.Condition[] { conGroup, conPane });
                    var childGroups = ret.BackStage.FindAll(System.Windows.Automation.TreeScope.Children, conGroupOrPane);

                    foreach (System.Windows.Automation.AutomationElement e in childGroups)
                    {
                        ret.Columns.Add(e);
                    }

                    return ret;
                }
                catch
                {
                    return new BackStageItem();
                }
            }

        }

        public class BackStageRect
        {
            public System.Windows.Rect BackStageView;
            public System.Windows.Rect BackStage;
            public readonly List<System.Windows.Rect> Columns = new List<System.Windows.Rect>();
        }

        public static BackStageRect GetBackStageRect(IntPtr hwnd)
        {
            return BackStageItem.GetBackStageItem(hwnd).ToBackStageRect();
        }

        public static void RefreshPrintTab(IntPtr hwnd)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                var item =BackStageItem.GetBackStageItem(hwnd);
                if (item.Columns.Count <= 0)
                {
                    return;
                }

                var con = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.AccessKeyProperty, "Alt, F, P, R");
                var btn = item.BackStage.FindFirst(System.Windows.Automation.TreeScope.Subtree, con);
                if (btn == null)
                {
                    return;
                }

                var ip = btn.GetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern) as System.Windows.Automation.InvokePattern;
                if (ip == null)
                {
                    return;
                }
                ip.Invoke();

                con = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ClassNameProperty, "#32770");
                System.Windows.Automation.AutomationElement dlg;

                IntPtr hwndDlg = IntPtr.Zero;
                do
                {
                    dlg = item.wnd.FindFirst(System.Windows.Automation.TreeScope.Subtree, con);
                } while (dlg == null || dlg.Current.NativeWindowHandle == 0);

                con = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Button);
                var buttons = dlg.FindAll(System.Windows.Automation.TreeScope.Subtree, con);
                foreach (System.Windows.Automation.AutomationElement b in buttons)
                {
                    System.Diagnostics.Debug.WriteLine(b.Current.Name + "\t" + b.Current.AccessKey);
                }

                con = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.AccessKeyProperty, "Alt+V");
                var con2 = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.AccessKeyProperty, "Alt+v");
                var orcon = new System.Windows.Automation.OrCondition(new System.Windows.Automation.Condition[] { con, con2 });
                var prev = dlg.FindFirst(System.Windows.Automation.TreeScope.Subtree, orcon);
                if (prev == null)
                {
                    return;
                }
                var ip2 = prev.GetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern) as System.Windows.Automation.InvokePattern;
                if (ip2 == null)
                {
                    return;
                }
                dlg.SetFocus();

                do
                {
                    ip2.Invoke();
                } while (Win32.IsWindowVisible(hwndDlg));

            });
        }
    }
}
