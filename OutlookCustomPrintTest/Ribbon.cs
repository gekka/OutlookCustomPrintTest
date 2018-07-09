
namespace OutlookCustomPrintTest
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using Office = Microsoft.Office.Core;


    [System.Runtime.InteropServices.ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Compose"
            || ribbonID == "Microsoft.Outlook.Mail.Read")
            {
                return GetResourceText("OutlookCustomPrintTest.Ribbon.xml");
            }
            return string.Empty;
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }


        #region 元の印刷タブでMailItemの中身をいじくってみる方法
        private Dictionary<Microsoft.Office.Interop.Outlook.Inspector, string> dicHTML = new Dictionary<Microsoft.Office.Interop.Outlook.Inspector, string>();

        public void onChangeToggle(Office.IRibbonControl control, bool isPressed)
        {
            var ins = Globals.ThisAddIn.Application.ActiveInspector() as Microsoft.Office.Interop.Outlook.Inspector;
            if (ins == null)
            {
                return;
            }

            var mail = ins.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;
            if (mail != null)
            {

                string originalHTML;
                if (dicHTML.ContainsKey(ins))
                {
                    originalHTML = dicHTML[ins];
                }
                else
                {
                    dicHTML.Add(ins, mail.HTMLBody);
                    originalHTML = mail.HTMLBody;
                }

                if (isPressed)
                {
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    dic.Add("Key1", "あいうえお");
                    dic.Add("key2", "かきくけこ");
                    mail.HTMLBody = CreateTestHtml(originalHTML, dic);
                }
                else
                {
                    mail.HTMLBody = originalHTML;
                }
            }
            IntPtr hwnd = BackStageTool.GetActiveWindowHwnd();
            BackStageTool.RefreshPrintTab(hwnd);
        }

        private static string CreateTestHtml(string html, Dictionary<string, string> dic)
        {
            try
            {
                System.Windows.Forms.WebBrowser wb = new System.Windows.Forms.WebBrowser();
                wb.Navigate("about:blank");

                var doc = wb.Document;
                doc.OpenNew(false);
                doc.Write(html);
                wb.Refresh();

                var div = doc.CreateElement("div");
                wb.Document.Body.InsertAdjacentElement(System.Windows.Forms.HtmlElementInsertionOrientation.AfterBegin, div);

                var table = doc.CreateElement("table");
                div.AppendChild(table);
                {
                    var tbody = doc.CreateElement("tbody");
                    table.AppendChild(tbody);
                    {
                        foreach (var keyv in dic)
                        {
                            var tr = doc.CreateElement("tr");
                            tbody.AppendChild(tr);

                            var td1 = doc.CreateElement("td");
                            td1.Style = "font-weight:bold";
                            td1.InnerText = keyv.Key;

                            var td2 = doc.CreateElement("td");
                            td2.InnerText = keyv.Value;
                            td2.Style = "color:red";
                            tr.AppendChild(td1);
                            tr.AppendChild(td2);
                        }
                    }
                }
                div.InsertAdjacentElement(System.Windows.Forms.HtmlElementInsertionOrientation.BeforeEnd, doc.CreateElement("hr"));

                return wb.Document.Body.OuterHtml;
            }
            catch (Exception ex)
            {
                return html;
            }
        }

        #endregion


        #region 自作タブで印刷ページを作ってすべて自分で印刷する方法

        public void backstage_Show(object context)
        {
            var scontext = System.Windows.Forms.WindowsFormsSynchronizationContext.Current;
            if (scontext == null)
            {
                scontext = new System.Windows.Forms.WindowsFormsSynchronizationContext();
                System.Windows.Forms.WindowsFormsSynchronizationContext.SetSynchronizationContext(scontext);
            }

            CustomPrintBackstageWindow.Create(scontext, this.ribbon);
        }

        public void backstage_Hide(object context)
        {
            CustomPrintBackstageWindow.OnHide(context);

            var ins = context as Microsoft.Office.Interop.Outlook.Inspector;

            dicHTML.Remove(ins);
        }

        public System.Drawing.Bitmap getPreviewImage(Office.IRibbonControl control)
        {
            return CustomPrintBackstageWindow.RequestBitmap(control);
        }

        public void onPrintCustom(Office.IRibbonControl control)
        {
            var wb = CustomPrintBackstageWindow.GetBackStageWindow(control.Context as Microsoft.Office.Interop.Outlook.Inspector);
            if (wb != null && wb.Page.CanPrint)
            {
                wb.Page.PrintPage();
            }
        }

        #endregion

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

    }

}

