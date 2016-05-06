using MsOfficeUtility.Common.Enums;
using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace MsOfficeUtility.Word
{
    /// <summary>
    /// http://www.e-iceblue.com/Tutorials.html
    /// </summary>
    public class SpireDocWrapper : IWord
    {
        private bool _disposed = false;
        private Document _doc;

        /// <summary>
        /// 
        /// </summary>
        public SpireDocWrapper()
        {
            _doc = new Document();
        }

        /// <summary>
        /// 
        /// </summary>
        ~SpireDocWrapper()
        {
            Dispose(false);//必须为false
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);//必须为true
            GC.SuppressFinalize(this);//通知垃圾回收机制不再调用终结器（析构器）
        }

        /// <summary>
        /// 
        /// </summary>
        public bool Open(string file)
        {
            _doc.LoadFromFile(file);
            return true;
        }

        /// <summary>
        /// http://www.e-iceblue.com/Tutorials/Spire.Doc/Spire.Doc-Program-Guide/Bookmark/How-to-insert-an-image-at-bookmark-in-word-documents.html
        /// </summary>
        public bool PilingPicture(string pictureFile, string bookmark, ImageWrapType wrapType = ImageWrapType.Front, ImageHorizontalAlignment horizontalAlignment = ImageHorizontalAlignment.Left)
        {
            Image image = Image.FromFile(pictureFile);

            //书签后插入图片
            BookmarksNavigator bn = new BookmarksNavigator(_doc);
            bn.MoveToBookmark(bookmark, true, true);
            var pic = bn.CurrentBookmark.BookmarkEnd.OwnerParagraph.AppendPicture(image);
            switch (wrapType)
            {
                case ImageWrapType.Behind:
                    pic.TextWrappingStyle = TextWrappingStyle.Behind;//衬于文字下方
                    break;
                case ImageWrapType.Front:
                    pic.TextWrappingStyle = TextWrappingStyle.InFrontOfText;//文字上方
                    break;
            }
            switch (horizontalAlignment)
            {
                case ImageHorizontalAlignment.Left:
                    pic.HorizontalAlignment = ShapeHorizontalAlignment.Left;
                    break;
                case ImageHorizontalAlignment.Right:
                    pic.HorizontalAlignment = ShapeHorizontalAlignment.Right;
                    break;
                case ImageHorizontalAlignment.Center:
                    pic.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                    break;
            }
            //pic.HorizontalOrigin = HorizontalOrigin.Margin;//找不到合适的对齐方式
            //pic.VerticalAlignment = ShapeVerticalAlignment.Center;//似乎都不起作用
            pic.VerticalOrigin = VerticalOrigin.Paragraph;//这个对齐了

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public bool PilingWord(Dictionary<string, string> data)
        {
            foreach (var d in data)
            {
                _doc.Replace(d.Key, d.Value, true, false);
            }
            return true;
        }

        /// <summary>
        /// http://www.e-iceblue.com/Tutorials/Spire.Doc/Spire.Doc-Program-Guide/Print-a-Word-Document-Programmatically-in-5-Steps.html
        /// </summary>
        public bool Print(string pages)
        {
            _doc.PrintDialog = new PrintDialog
            {
                AllowCurrentPage = true,
                AllowSomePages = true,
                //PrinterSettings = new PrinterSettings
                //{
                //    MaximumPage = _doc.PageCount
                //}
            };

            PrintDocument printDoc = _doc.PrintDocument;

            //printDoc.PrintController = new StandardPrintController();//without showing print processing dialog

            //if (dialog.ShowDialog() == DialogResult.OK)
            {
                printDoc.Print();
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public bool Save()
        {
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        public bool SaveAs(string file)
        {
            _doc.SaveToFile(file);

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }
            if (disposing)
            {
                //清理托管资源
                _doc.Close();
                _doc.Dispose();
            }
            //清理非托管资源

            _disposed = true;//让类型知道自己已经被释放
        }
    }
}
