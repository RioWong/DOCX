using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.format
{
    /// <summary>
    /// 格式控制函数
    /// </summary>
    public class formatHelper
    {
        /// <summary>
        /// 设置段落的格式
        /// </summary>
        /// <param name="fontFamily">字体对象</param>
        /// <param name="fontsize">字体大小</param>
        /// <param name="color">颜色</param>
        /// <param name="isbold">加粗</param>
        /// <returns></returns>
        public static Formatting SetParagraphFormat(FontFamily fontFamily, int fontsize, Color color, bool isbold = false)
        {
            Formatting f = new Formatting();
            f.Bold = isbold;
            f.FontFamily = fontFamily;
            f.FontColor = color;
            f.Size = fontsize;

            return f;
        }
    }
}
