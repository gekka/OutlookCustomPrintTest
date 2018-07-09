using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Office = Microsoft.Office.Core;
namespace OutlookCustomPrintTest
{
    /// <summary>BackstageViewの印刷タブ用</summary>
    class CustomPrintBackstageWindow : System.Windows.Forms.NativeWindow
    {
        private static Dictionary<IntPtr, CustomPrintBackstageWindow> dic = new Dictionary<IntPtr, CustomPrintBackstageWindow>();

        private const string IMAGECONTROL_PRINTIMAGE = "previewImage";

        private Microsoft.Office.Interop.Outlook.Inspector inspector;
        private Microsoft.Office.Interop.Outlook.MailItem mail;
        private System.Threading.SynchronizationContext scontext;
        private Office.IRibbonUI ribbonUI;

        public Page Page { get { return _Page; } }
        private Page _Page = new Page();

        public static void Create(System.Threading.SynchronizationContext scontext, Office.IRibbonUI ribbonUI)
        {
            scontext.Post((o) =>
            {
                var hwnd = BackStageTool.GetActiveWindowHwnd();
                if (hwnd != IntPtr.Zero)
                {
                    lock (dic)
                    {
                        CustomPrintBackstageWindow w;
                        if (!dic.TryGetValue(hwnd, out w))
                        {
                            w = new CustomPrintBackstageWindow();
                            w.AssignHandle(BackStageTool.GetActiveWindowHwnd());
                            w.scontext = scontext;
                            w.ribbonUI = ribbonUI;
                            w.inspector = Globals.ThisAddIn.Application.ActiveInspector() as Microsoft.Office.Interop.Outlook.Inspector;
                            w.mail = w.inspector.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;
                            if (w.mail != null)
                            {
                                w.Page.Html = w.mail.HTMLBody;
                            }

                            dic.Add(hwnd, w);
                        }
                    }
                }
            }, null);
        }

        public static CustomPrintBackstageWindow GetBackStageWindow(Microsoft.Office.Interop.Outlook.Inspector ins)
        {
            if (ins != null)
            {
                lock (dic)
                {
                    var w = dic.Values.FirstOrDefault(_ => _.inspector == ins);
                    if (w != null)
                    {
                        return w;
                    }
                }
            }
            return null;
        }

        public override void ReleaseHandle()
        {
            lock (dic)
            {
                dic.Remove(this.Handle);
            }
            base.ReleaseHandle();
        }

        public static void OnHide(object context)
        {
            var wb = GetBackStageWindow(context as Microsoft.Office.Interop.Outlook.Inspector);
            if (wb != null)
            {
                wb.ReleaseHandle();
            }
        }

        public static System.Drawing.Bitmap RequestBitmap(Office.IRibbonControl control)
        {
            var ins = control.Context as Microsoft.Office.Interop.Outlook.Inspector;
            if (ins == null)
            {
                return null;
            }

            lock (dic)
            {
                var w = GetBackStageWindow(ins);
                if (w != null)
                {
                    if (w.Page.Bmp != null)
                    {
                        return w.Page.Bmp;
                    }

                    w.scontext.Post((o) =>
                    {
                        w.CreateBitmap();
                    }, null);
                }
            }
            return null;
        }

        private void CreateBitmap()
        {
            if (!this.Page.CanPrint)
            {
                return;
            }
            var rects = BackStageTool.GetBackStageRect(this.Handle);
            if (rects.Columns.Count < 2)
            {
                return;
            }

            double w = Math.Max(0, rects.Columns[1].Width - 20);
            double h = Math.Max(0, rects.BackStageView.Height - 20);


            this.Page.CreateBitmap(w, h, this.Handle);
            this.ribbonUI.InvalidateControl(IMAGECONTROL_PRINTIMAGE);
        }

        protected void OnSizeChanged()
        {
            if (this.ribbonUI != null)
            {
                this.Page.CrearBmp();
                ribbonUI.InvalidateControl(IMAGECONTROL_PRINTIMAGE);
            }
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            base.WndProc(ref m);

            switch (m.Msg)
            {
            case 5:
                OnSizeChanged();
                break;
            case 10:
                this.ReleaseHandle();
                break;
            }
        }


     
    }


   
}
