using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using InteropWord = Microsoft.Office.Interop.Word;

namespace Hang.Net.Office.Word
{
    /// <summary>
    /// Microsoft.Office.Interop.Word公共操作
    /// </summary>
    public static class MsWordUtility
    {
        /// <summary>
        /// 打印一份Doc文档
        /// </summary>
        /// <param name="docFile"></param>
        public static void Print(string docFile)
        {
            object wordFile = docFile;

            InteropWord.Application app = null;
            InteropWord.Document doc = null;

            try
            {
                app = new InteropWord.Application();
                app.Visible = false;
                app.DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone;

                doc = app.Documents.Open(ref wordFile);

                doc.PrintOut();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
                if (app != null)
                {
                    app.Quit();
                }
            }
        }

        /// <summary>
        /// Convert doc to image
        /// </summary>
        /// <param name="docFile"></param>
        /// <param name="imagePath"></param>
        public static void ToImage(string docFile, string imagePath)
        {
            string docFileName = docFile.Remove(0, docFile.LastIndexOf("\\") + 1);
            object wordFile = docFile;

            InteropWord.Application app = null;
            InteropWord.Document doc = null;

            try
            {
                app = new InteropWord.Application();
                app.Visible = false;
                app.DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone;

                doc = app.Documents.Open(ref wordFile);

                doc.ShowGrammaticalErrors = false;
                doc.ShowRevisions = false;
                doc.ShowSpellingErrors = false;

                //Opens the word document and fetch each page and converts to image
                foreach (InteropWord.Window window in doc.Windows)
                {
                    foreach (InteropWord.Pane pane in window.Panes)
                    {
                        for (var i = 1; i <= pane.Pages.Count; i++)
                        {
                            var page = pane.Pages[i];
                            var bits = page.EnhMetaFileBits;
                            var target = Path.Combine(imagePath + "\\" + docFileName.Split('.')[0], string.Format("{1}_page_{0}", i, imagePath.Split('.')[0]));

                            try
                            {
                                using (var ms = new MemoryStream((byte[])(bits)))
                                {
                                    var image = Image.FromStream(ms);
                                    var pngTarget = Path.ChangeExtension(target, "png");
                                    image.Save(pngTarget, ImageFormat.Png);
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
                if (app != null)
                {
                    app.Quit();
                }
            }
        }
    }
}
