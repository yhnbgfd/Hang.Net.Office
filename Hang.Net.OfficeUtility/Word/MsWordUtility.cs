using System;
using InteropWord = Microsoft.Office.Interop.Word;

namespace Hang.Net.OfficeUtility.Word
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
    }
}
