using Hang.Net.OfficeUtility.Common.Enums;
using System;
using System.Collections.Generic;

namespace Hang.Net.OfficeUtility.Word
{
    /// <summary>
    /// 
    /// </summary>
    public interface IWord : IDisposable
    {
        /// <summary>
        /// 打开
        /// </summary>
        /// <param name="fileName"></param>
        bool Open(string fileName);
        /// <summary>
        /// 保存
        /// </summary>
        bool Save();
        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="fileName"></param>
        bool SaveAs(string fileName);
        /// <summary>
        /// 查找替换文本
        /// </summary>
        /// <param name="data"></param>
        bool PilingWord(Dictionary<string, string> data);
        /// <summary>
        /// 标签处插入图片
        /// </summary>
        /// <param name="pictureFile"></param>
        /// <param name="bookmark"></param>
        /// <param name="wrapType"></param>
        /// <param name="horizontalAlignment"></param>
        /// <returns></returns>
        bool PilingPicture(string pictureFile, string bookmark, ImageWrapType wrapType = ImageWrapType.Front, ImageHorizontalAlignment horizontalAlignment = ImageHorizontalAlignment.Left);
        /// <summary>
        /// 打印
        /// </summary>
        bool Print(string pages);
    }
}
