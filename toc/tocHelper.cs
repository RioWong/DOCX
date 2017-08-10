using CreateWord.log;
using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.toc
{
    public class tocHelper
    {
        /// <summary>
        /// 根据书签地址生成目录
        /// </summary>
        /// <param name="document"></param>
        /// <param name="markname"></param>
        static public void AddtocByBookmark(DocX document, string markname)
        {
            try
            {
                document.InsertTableOfContents(document.Bookmarks[markname].Paragraph, "目录", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2", 5);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tocHelper), ex);
            }
        }
    }
}
