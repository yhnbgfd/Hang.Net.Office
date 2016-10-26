using Hang.Net.Office.Word;
using NLog;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Demo
{
    public class TestWord
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Dictionary<string, string> _dictData;

        public TestWord()
        {
            _dictData = new Dictionary<string, string>();
            _dictData.Add("{文本替换测试1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            _dictData.Add("{文本替换测试2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            _dictData.Add("{文本替换测试3}", Guid.NewGuid().ToString());
            _dictData.Add("{文本替换测试4}", Guid.NewGuid().ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        public void TestSpireDoc()
        {
            //try
            //{
            //    using (IWord w = new SpireDocWrapper())
            //    {
            //        w.Open(AppDomain.CurrentDomain.BaseDirectory + @"Resources\Test.docx");
            //        w.PilingWord(_dictData);
            //        w.PilingPicture(AppDomain.CurrentDomain.BaseDirectory + @"Resources\IDCard.bmp", "测试用书签1");
            //        //w.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"SpireDoc.docx");
            //        w.Print(null);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        /// <summary>
        /// 
        /// </summary>
        public void TestMicrosoftWord(string extend = ".doc")
        {
            try
            {
                _logger.Trace("1 Start Test TestMicrosoftWord");

                using (IWord w = new MSWordWrapper())
                {
                    _logger.Trace("2 new MSWordWrapper()");

                    w.Open(AppDomain.CurrentDomain.BaseDirectory + @"Resources\Test" + extend);

                    _logger.Trace("3 Open");

                    w.PilingWord(_dictData);

                    _logger.Trace("4 PilingWord");

                    w.PilingPicture(AppDomain.CurrentDomain.BaseDirectory + @"Resources\IDCard.bmp", "测试用书签1");

                    _logger.Trace("5 PilingPicture");

                    w.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"MicrosoftWord" + extend);

                    _logger.Trace("6 SaveAs");

                    w.Print("1,3-4");
                }
                //MsWordUtility.Print(AppDomain.CurrentDomain.BaseDirectory + @"MicrosoftWord" + extend);

                _logger.Trace("7 Print");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

    }
}
