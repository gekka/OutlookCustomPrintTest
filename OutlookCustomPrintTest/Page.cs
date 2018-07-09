namespace OutlookCustomPrintTest
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>HTMLを指定ページサイズの画像化</summary>
    public class Page
    {
        private OutlookCustomPrintTest.HTMLCapture capture = new HTMLCapture();

        public Page()
        {
            Width = 210;
            Height = 297;
            MarginL = 5;
            MarginT = 5;
            MarginR = 5;
            MarginB = 5;

            capture.Ready += capture_Ready;
        }

        void capture_Ready(object sender, EventArgs e)
        {
            CreateMetaFile();
        }

        public string Html { get; set; }

        public float Width { get; set; }
        public float Height { get; set; }
        public float MarginL { get; set; }
        public float MarginT { get; set; }
        public float MarginR { get; set; }
        public float MarginB { get; set; }

        public void CrearBmp()
        {
            this._Bmp = null;
        }

        public System.Drawing.Bitmap Bmp
        {
            get { return _Bmp; }
        }
        private System.Drawing.Bitmap _Bmp;
        private System.Drawing.Imaging.Metafile meta;

        /// <summary>Bitmap,印刷共通で使うメタファイルを作る</summary>
        public void CreateMetaFile()
        {
            this.meta = this.capture.ToMetaFile(this.Html, Width, Height, MarginL, MarginT, MarginR, MarginB);
        }
        private void SaveMetaFile(System.IO.Stream output, System.Drawing.Imaging.Metafile meta, IntPtr handle)
        {
            var graHandle = System.Drawing.Graphics.FromHwnd(handle);
            var hdc = graHandle.GetHdc();
            var metaClone = new System.Drawing.Imaging.Metafile(output, hdc);
            System.Drawing.Graphics graMeta;
            graMeta = System.Drawing.Graphics.FromImage(metaClone);
            graMeta.DrawImage(meta, System.Drawing.Point.Empty);
            graMeta.Dispose();
            graHandle.ReleaseHdc(hdc);
            graHandle.Dispose();
        }

        /// <summary>指定サイズに収まるbitmapを作る</summary>
        /// <param name="w">作りたいbmpの幅</param>
        /// <param name="h">作りたいbmpの高さ</param>
        public void CreateBitmap(double w, double h, IntPtr hwnd)
        {
            this._Bmp = null;
            if (this.meta == null)
            {
                return;
            }

            double aspect = Width / Height;
            if (w / h < aspect)
            {
                h = w / aspect;
            }
            else
            {
                w = h * aspect;
            }

            System.Drawing.Rectangle rect = new System.Drawing.Rectangle();
            rect.X = (int)MarginL;
            rect.Y = (int)MarginT;
            rect.Width = (int)w;
            rect.Height = (int)h;


            if (rect.Width == 0 & rect.Height == 0)
            {

                return;
            }
            this._Bmp = new System.Drawing.Bitmap((int)w, (int)h);

            System.Drawing.PointF dpi = new System.Drawing.PointF();
            using (var g = System.Drawing.Graphics.FromHwnd(hwnd))
            {
                dpi.X = g.DpiX;
                dpi.Y = g.DpiY;
            }
            
            double bmpDpiX = Bmp.Width / (Width / 25.4 * dpi.X) * dpi.X;
            double bmpDpiY = Bmp.Height / (Height / 25.4 * dpi.Y) * dpi.Y;
            Bmp.SetResolution((float)bmpDpiX, (float)bmpDpiY);

            using (var g = System.Drawing.Graphics.FromImage(Bmp))
            {
                g.PageUnit = System.Drawing.GraphicsUnit.Millimeter;
                g.FillRectangle(System.Drawing.Brushes.White, new System.Drawing.RectangleF(0, 0, (float)Width, (float)Height));
                g.DrawRectangle(System.Drawing.Pens.Black, new System.Drawing.Rectangle(0, 0, (int)Width-1, (int)Height-1));

                double mw = meta.Width / meta.HorizontalResolution * 25.4;
                double mh = mw * h/w;
                float innerW = (Width - MarginL - MarginR);
                float innerH = (Height - MarginT - MarginB);
                var rd = new System.Drawing.RectangleF(MarginL, MarginT, innerW, innerH);
                var rs = new System.Drawing.RectangleF(0, 0, (float)mw, (float)mh);

                g.DrawImage(this.meta, rd, rs, System.Drawing.GraphicsUnit.Millimeter);
            }
        }

        /// <summary>印刷もしくはBitmapを作れるか</summary>
        public bool CanPrint
        {
            get { return meta != null; }
        }

        /// <summary>印刷する</summary>
        public void PrintPage()
        {
            if (!CanPrint)
            {
                throw new InvalidOperationException();
            }

            System.Windows.Forms.PrintDialog dlg = new System.Windows.Forms.PrintDialog();
            dlg.Document = new System.Drawing.Printing.PrintDocument();

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dlg.Document.PrintPage += Document_PrintPage;
                dlg.Document.Print();
            }


        }
        private void Document_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            var g = e.Graphics;
            g.DrawImage(meta, System.Drawing.Point.Empty);
            e.HasMorePages = false;
        }
    }
}
