namespace OutlookCustomPrintTest
{
    using System;
    using System.Drawing;
    using System.Windows.Forms;

    public class HTMLCapture
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        private static extern int OleDraw(IntPtr pUnk, int dwAspect, IntPtr hdcDraw, ref RECT lprcBounds);

        private CustomWebBrowser wb;

        public HTMLCapture()
        {
            wb = new OutlookCustomPrintTest.CustomWebBrowser();
            wb.DocumentCompleted += wb_DocumentCompleted;
            wb.ScrollBarsEnabled = false;
            wb.ScriptErrorsSuppressed = true;
            wb.Navigate("about:blank");
            wb.Refresh();
        }

        void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.wb.DocumentCompleted -= wb_DocumentCompleted;
            CanCapture = true;
            var r = Ready;
            if (r != null)
            {
                r(this, EventArgs.Empty);
            }
        }
        public event EventHandler Ready;
        public bool CanCapture { get; private set; }

        public System.Drawing.Imaging.Metafile ToMetaFile_(string html, double width, double height, double marginL, double marginT, double marginR, double marginB)
        {
            double w = width - marginL - marginR;
            double h = height - marginT - marginB;
            string pageSizeStyle = string.Format("width: {0}mm;height: {1}mm;margin-left: 0px;margin-right: 0px;", w, h);
            wb.Document.OpenNew(false);
            wb.Document.Write(html);
            wb.Refresh();

            wb.Document.Body.Style = pageSizeStyle;
            wb.Refresh();

            wb.Width = wb.Document.Body.ScrollRectangle.Width;
            wb.Height = wb.Document.Body.ScrollRectangle.Height;

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            using (var graHandle = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                double dpix = graHandle.DpiX;
                double dpiy = graHandle.DpiY;
                var hdc = graHandle.GetHdc();
                System.Drawing.Imaging.Metafile meta
                    = new System.Drawing.Imaging.Metafile(ms, hdc, new System.Drawing.RectangleF(0, 0, (float)w,(float) h), System.Drawing.Imaging.MetafileFrameUnit.Millimeter);
                using (var g = Graphics.FromImage(meta))
                {
                }
                double hres = meta.HorizontalResolution;
                double vres = meta.VerticalResolution;

                meta = new System.Drawing.Imaging.Metafile(ms, hdc, new System.Drawing.RectangleF(0, 0, (float)w, (float)h), System.Drawing.Imaging.MetafileFrameUnit.Millimeter);
                using (var g = Graphics.FromImage(meta))
                {
                    double scale = hres / 25.4;
                    RECT r = new RECT();
                    r.Left = 0;
                    r.Top = 0;
                    r.Right = (int)w;
                    r.Bottom = (int)((double)w * (double)wb.Document.Body.ScrollRectangle.Height / (double)wb.Document.Body.ScrollRectangle.Width);
                    r.Right *= 10;
                    //g.PageUnit = GraphicsUnit.Millimeter;
                    try
                    {
                        var hdc2 = g.GetHdc();
                        try
                        {
                            var unknown = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(wb.ActiveXInstance);
                            try
                            {
                                OleDraw(unknown, (int)System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT, hdc2, ref r);
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.Release(unknown);
                            }
                        }
                        finally
                        {
                            g.ReleaseHdc(hdc2);
                        }
                        for (int x = 0; x <= width; x += 10)
                        {
                            g.DrawLine(System.Drawing.Pens.Red, new System.Drawing.Point(x, 0), new System.Drawing.Point(x, 297));
                        }
                    }
                    finally
                    {
                        graHandle.ReleaseHdc(hdc);
                    }
                }
                meta.Dispose();
            }

            ms.Position = 0;
            using (System.IO.FileStream fs = System.IO.File.OpenWrite("j:\\test.emf"))
            {
                ms.WriteTo(fs);
            }
            ms.Position = 0;

            return System.Drawing.Imaging.Metafile.FromStream(ms) as System.Drawing.Imaging.Metafile; 
        }

        public System.Drawing.Imaging.Metafile ToMetaFile(string html, double width, double height, double marginL, double marginT, double marginR, double marginB)
        {
            double w = width - marginL - marginR;
            double h = height - marginT - marginB;
            string pageSizeStyle = string.Format("width: {0}mm;height: {1}mm;margin-left: 0px;margin-right: 0px;", w, h);
            wb.Document.OpenNew(false);
            wb.Document.Write(html);
            wb.Refresh();

            wb.Document.Body.Style = pageSizeStyle;
            wb.Refresh();

            wb.Width = wb.Document.Body.ScrollRectangle.Width;
            wb.Height = wb.Document.Body.ScrollRectangle.Height;

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            using (var graHandle = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                double dpix = graHandle.DpiX;
                double dpiy = graHandle.DpiY;
                var hdc = graHandle.GetHdc();
                System.Drawing.Imaging.Metafile meta
                    = new System.Drawing.Imaging.Metafile(ms, hdc, new System.Drawing.RectangleF(0, 0, (float)w, (float)h), System.Drawing.Imaging.MetafileFrameUnit.Millimeter);

                using (var g = Graphics.FromImage(meta))
                {
                    g.PageUnit = GraphicsUnit.Millimeter;

                    var xd = dpix / meta.PhysicalDimension.Width;
                    var yd = dpiy / meta.PhysicalDimension.Height;


                    RECT r = new RECT();
                    r.Left = 0;
                    r.Top = 0;
                    r.Right = (int)(w*xd);
                    r.Bottom = (int)((double)w * (double)wb.Document.Body.ScrollRectangle.Height / (double)wb.Document.Body.ScrollRectangle.Width * yd);
                  
                    
                    try
                    {
                        var hdc2 = g.GetHdc();
                        try
                        {
                            var unknown = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(wb.ActiveXInstance);
                            try
                            {
                                OleDraw(unknown, (int)System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT, hdc2, ref r);
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.Release(unknown);
                            }
                        }
                        finally
                        {
                            g.ReleaseHdc(hdc2);
                        }


                        var state = g.Save();
                        g.TranslateTransform(0, 100);
                        g.RotateTransform(-45);
                        
                        g.DrawString("てすと", new Font("Meiryo", 100), Brushes.Red, 0, 0);
                        g.Restore(state);
                    }
                    finally
                    {
                        graHandle.ReleaseHdc(hdc);
                    }
                }
                meta.Dispose();
            }

            ms.Position = 0;

            return System.Drawing.Imaging.Metafile.FromStream(ms) as System.Drawing.Imaging.Metafile;
        }
    }
}