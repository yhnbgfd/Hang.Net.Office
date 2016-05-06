using MsOfficeUtility.Common.Enums;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using InteropWord = Microsoft.Office.Interop.Word;

namespace MsOfficeUtility.Word
{
    /// <summary>
    /// Microsoft.Office.Interop.Word封装操作
    /// </summary>
    public class MSWordWrapper : IWord
    {
        private bool _disposed = false;
        private object _missing = Missing.Value;

        private InteropWord.Application _app = null;
        private InteropWord.Document _doc = null;

        /// <summary>
        /// 析构函数
        /// </summary>
        ~MSWordWrapper()
        {
            Dispose(false);//必须为false
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);//必须为true
            GC.SuppressFinalize(this);//通知垃圾回收机制不再调用终结器（析构器）
        }

        /// <summary>
        /// Dispose
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }
            if (disposing)//清理托管资源
            {
                if (_doc != null)
                {
                    _doc.Close();
                }
                if (_app != null)
                {
                    _app.Quit();
                }
            }
            // 清理非托管资源

            _disposed = true;//让类型知道自己已经被释放
        }

        /// <summary>
        /// 打开Doc文件
        /// </summary>
        /// <param name="templateFile"></param>
        public bool Open(string templateFile)
        {
            object wordFile = templateFile;

            try
            {
                _app = new InteropWord.Application();
                _app.DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone;

                _doc = _app.Documents.Open(ref wordFile);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        public bool Save()
        {
            try
            {
                _doc.Save();
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 另存为文件
        /// </summary>
        /// <param name="saveFile"></param>
        public bool SaveAs(string saveFile)
        {
            try
            {
                _doc.SaveAs(saveFile);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 查找替换文本
        /// </summary>
        /// <param name="data"></param>
        public bool PilingWord(Dictionary<string, string> data)
        {
            object replace = InteropWord.WdReplace.wdReplaceAll;

            try
            {
                foreach (var d in data)
                {
                    _app.Selection.Find.Replacement.ClearFormatting();
                    _app.Selection.Find.ClearFormatting();
                    _app.Selection.Find.Text = d.Key;//需要被替换的文本
                    _app.Selection.Find.Replacement.Text = d.Value;//替换文本 

                    _app.Selection.Find.Execute(
                            ref _missing, ref _missing, ref _missing, ref _missing,
                            ref _missing, ref _missing, ref _missing, ref _missing,
                            ref _missing, ref _missing, ref replace,
                            ref _missing, ref _missing, ref _missing, ref _missing);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 定位书签并在书签位置插入图片(衬于文字下方)
        /// </summary>
        public bool PilingPicture(string pictureFile, string bookMark, ImageWrapType wrapType = ImageWrapType.Front, ImageHorizontalAlignment horizontalAlignment = ImageHorizontalAlignment.Left)
        {
            bool ret = false;

            try
            {
                if (_doc.Bookmarks.Exists(bookMark) == true)
                {
                    Image img = Image.FromFile(pictureFile);

                    //_doc.Bookmarks[bookMark].Select();
                    //InteropWord.InlineShape inlineShape = _app.ActiveDocument.InlineShapes.AddPicture(pictureFile, false, true, anchor);//InlineShapes这种方式无法定位
                    //inlineShape.Width = img.Width;
                    //inlineShape.Height = img.Height;
                    //inlineShape.ScaleHeight = 100f;
                    //inlineShape.ScaleWidth = 100f;

                    object anchor = _doc.Bookmarks[bookMark].Range;

                    InteropWord.Shape shape = _app.ActiveDocument.Shapes.AddPicture(pictureFile, false, true, 0f, 0f, img.Width * 1f, img.Height * 1f, anchor);
                    shape.ScaleHeight(1f, Microsoft.Office.Core.MsoTriState.msoTrue);
                    shape.ScaleWidth(1f, Microsoft.Office.Core.MsoTriState.msoTrue);
                    //https://msdn.microsoft.com/zh-cn/library/microsoft.office.interop.word.shape.scalewidth%28v=office.11%29.aspx
                    //https://msdn.microsoft.com/zh-cn/library/microsoft.office.core.msotristate%28v=office.11%29.aspx

                    //shape.LeftRelative = 0f;//被重置了所以再设一遍
                    //shape.TopRelative = 0f;//XP,2003下会报错.去掉之后,doc格式的图片位置正常

                    switch (wrapType)
                    {
                        case ImageWrapType.Behind:
                            //inlineShape.ConvertToShape().WrapFormat.Type = InteropWord.WdWrapType.wdWrapBehind;//衬于文字下方
                            shape.WrapFormat.Type = InteropWord.WdWrapType.wdWrapBehind;
                            break;
                        case ImageWrapType.Front://默认衬于文字上方
                            //inlineShape.ConvertToShape().WrapFormat.Type = InteropWord.WdWrapType.wdWrapFront;//衬于文字上方
                            shape.WrapFormat.Type = InteropWord.WdWrapType.wdWrapFront;
                            break;
                    }

                    switch (horizontalAlignment)
                    {
                        case ImageHorizontalAlignment.Center:
                            break;
                        case ImageHorizontalAlignment.Left:
                            break;
                        case ImageHorizontalAlignment.Right:
                            break;
                    }

                    ret = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ret;
        }

        /// <summary>
        /// 调用Doc打印
        /// </summary>
        public bool Print(string pages)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(pages))
                {
                    _doc.PrintOut();
                }
                else
                {
                    _doc.PrintOut(Range: InteropWord.WdPrintOutRange.wdPrintRangeOfPages, Pages: pages);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
