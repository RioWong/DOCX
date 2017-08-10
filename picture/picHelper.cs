using CreateWord.log;
using Novacode;
using RealEstate.Logic;
using System.IO;

namespace CreateWord.picture
{

    class picHelper
    {
        ///// <summary>
        ///// 将图片插入到指定的书签位置
        ///// </summary>
        ///// <param name="document">操作的文档</param>
        ///// <param name="BMname">书签的名字</param>
        ///// <param name="picPath">图片的路径</param>
        //public static void insertBybookmark(DocX document,string BMname,string picPath)
        //{
        //    
        //}

        /// <summary>
        /// 把图片插入到段落
        /// </summary>
        /// <param name="p"></param>
        /// <param name="picPath"></param>
        public static void insert(DocX document, Paragraph p, string picPath,int height,int width)
        {

            try
            {
                //Image image = document.AddImage(picPath);

                Stream urlStream = PathManager.getSingleton().PathProviderAsPrimary.OpenReadStream(picPath, true);
                if( urlStream != null )
                {
                    MemoryStream ms = new MemoryStream();
                    urlStream.CopyTo(ms);
                    ms.Position = 0;
                    urlStream.Close();

                    Image image = document.AddImage(ms);
                    Picture picture = image.CreatePicture();
                    picture.Height = height;
                    picture.Width = width;
                    p.AppendPicture(picture).Alignment = Alignment.center;
                }

            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(picHelper), ex);
            }
        }

        /// <summary>
        /// 根据路径生成一个图片，返回图片
        /// </summary>
        /// <param name="document"></param>
        /// <param name="picPath"></param>
        /// <param name="height">1厘米约等于38像素</param>
        /// <param name="width">1厘米约等于38像素</param>
        /// <returns></returns>
        public static Picture getPic(DocX document, string picPath, int height, int width)
        {
            Picture p = null;
            try
            {
                //Image image = document.AddImage(picPath);

                Stream urlStream = PathManager.getSingleton().PathProviderAsPrimary.OpenReadStream(picPath, true);

                //Stream urlStream = System.Net.WebRequest.Create("http://fdzy.njnu.edu.cn:5791/%E5%9B%BD%E7%BD%91%E6%B1%9F%E8%8B%8F%E7%9C%81%E7%94%B5%E5%8A%9B%E5%85%AC%E5%8F%B8/%E5%8D%95%E4%BD%8D%E4%BB%8B%E7%BB%8D%E5%9B%BE.jpg").GetResponse().GetResponseStream();
                if (urlStream == null) return p;
                MemoryStream ms = new MemoryStream();
                urlStream.CopyTo(ms);
                ms.Position = 0;
                urlStream.Close();

                Image image = document.AddImage(ms);
                p = image.CreatePicture();
                p.Height = height;
                p.Width = width;
            }            
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(picHelper), ex);
            }
            return p;
        }

    }
}
