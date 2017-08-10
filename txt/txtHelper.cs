using CreateWord.log;
using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.txt
{

    public class txtHelper
    {
        /// <summary>
        /// 读取txt，按行读取
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> txtLines(string path)
        {
            List<string> lstStr = new List<string>();
            try
            {
                Stream urlStream = PathManager.getSingleton().PathProviderAsPrimary.OpenReadStream(path, true);
                if (urlStream == null) return lstStr;
                MemoryStream mStream = new MemoryStream();
                urlStream.CopyTo(mStream);
                mStream.Position = 0;
                urlStream.Close();

                //Stream stream = System.Net.WebRequest.Create(path).GetResponse().GetResponseStream();
                //MemoryStream mStream = new MemoryStream();
                //stream.CopyTo(mStream);
                byte[] bytes = mStream.ToArray();
                string text = Encoding.Default.GetString(bytes);
                string[] lines = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                //string[] lines = System.IO.File.ReadAllLines(path, Encoding.Default);
                foreach (string line in lines)
                {
                    lstStr.Add(line);
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(txtHelper), ex);
            }

            return lstStr;
        }
        /// <summary>
        /// 读取txt文件，并以string形式返回里面所有内容
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string readtxt(string path)
        {
            string text = "";
            try
            {
                Stream urlStream = PathManager.getSingleton().PathProviderAsPrimary.OpenReadStream(path, true);
                if (urlStream == null) return text;

                MemoryStream mStream = new MemoryStream();
                urlStream.CopyTo(mStream);
                mStream.Position = 0;
                urlStream.Close();

                //Stream stream = System.Net.WebRequest.Create(path).GetResponse().GetResponseStream();
                //MemoryStream mStream = new MemoryStream();
                //stream.CopyTo(mStream);
                byte[] bytes = mStream.ToArray();
                text = Encoding.Default.GetString(bytes);
               
                //text = System.IO.File.ReadAllText(path, Encoding.Default);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(txtHelper), ex);
            }
            return text;
        }
    }
}
