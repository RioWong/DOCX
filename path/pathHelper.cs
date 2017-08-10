using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace CreateWord.path
{
    /// <summary>
    /// 统一管理整个项目的路径
    /// </summary>
    public class pathHelper
    {
        /// <summary>
        /// 得到文档输出路径
        /// </summary>
        /// <returns></returns>
        public static string GetOutputPath(int CompanyID)
        {
            //return Environment.CurrentDirectory + "\\汇编文档.docx";
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            parentPath = parentPath.Replace("/", "\\");
            return HttpRuntime.AppDomainAppPath + parentPath + "汇编文档.docx";
        }

        /// <summary>
        /// 根据公司ID获取“单位介绍.txt”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <returns></returns>
        public static string GetIntrotxtPath(int CompanyID)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            parentPath = parentPath.Replace("/", "\\");
            return HttpRuntime.AppDomainAppPath + parentPath + "单位介绍.txt";
        }

        /// <summary>
        /// 根据公司ID获取“单位介绍图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <returns></returns>
        public static string GetIntropicPath(int CompanyID)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            parentPath = parentPath.Replace("/", "\\");
            return HttpRuntime.AppDomainAppPath + parentPath + "单位介绍图.jpg";
        }

        /// <summary>
        /// 根据公司ID获取“位置分布图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <returns></returns>
        public static string GetLocpicPath(int CompanyID)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            parentPath = parentPath.Replace("/", "\\");
            return HttpRuntime.AppDomainAppPath + parentPath + "位置分布图.jpg";
        }

        /// <summary>
        /// 根据公司ID和宗地名称获取“宗地图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <returns></returns>
        public static string GetParcelpicPath(int CompanyID, string parcelname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\宗地图.jpg";
        }

        /// <summary>
        /// 根据公司ID和宗地名称获取“土地证.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <returns></returns>
        public static string GetLandusepicPath(int CompanyID, string parcelname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\土地证.jpg";
        }

        /// <summary>
        /// 根据公司ID和宗地名称获取“总平面分布图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <returns></returns>
        public static string GetZPMFBpicPath(int CompanyID, string parcelname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\总平面分布图.jpg";
        }

        /// <summary>
        /// 根据公司ID、宗地名称和鸟瞰图的方向获取鸟瞰图的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <param name="direction">正射、前、后、左、右</param>
        /// <returns></returns>
        public static string GetAerialviewpicPath(int CompanyID, string parcelname, string direction)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\鸟瞰图\\" + direction + ".JPG";
        }

        /// <summary>
        /// 根据公司ID、宗地名称和建筑名称获取“外墙实景图.JPG”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <param name="buildingname"></param>
        /// <returns></returns>
        public static string GetVirtualmapPath(int CompanyID, string parcelname, string buildingname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\" + buildingname + "\\外墙实景图.JPG";
        }

        /// <summary>
        /// 根据公司ID、宗地名称和建筑名称获取“列表.txt”（分层分户平面图的信息）的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <param name="buildingname"></param>
        /// <returns></returns>
        public static string GetPlantxtPath(int CompanyID, string parcelname, string buildingname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\" + buildingname + "\\列表.txt";
        }

        /// <summary>
        /// 根据公司ID、宗地名称、建筑名称和图片名字获取分层分户平面图的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <param name="buildingname"></param>
        /// <param name="picname"></param>
        /// <returns></returns>
        public static string GetPlanpicPath(int CompanyID, string parcelname, string buildingname, string picname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\" + buildingname + "\\" + picname + ".jpg";
        }

        /// <summary>
        /// 根据公司ID和宗地名称获取“管线图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <param name="parcelname"></param>
        /// <returns></returns>
        public static string GetPipelinepicPath(int CompanyID, string parcelname)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            return HttpRuntime.AppDomainAppPath + parentPath + parcelname + "\\管线图.jpg";
        }

        /// <summary>
        /// 根据公司ID获取“组织结构图.jpg”的路径
        /// </summary>
        /// <param name="CompanyID"></param>
        /// <returns></returns>
        public static string GetOrganizationpicPath(int CompanyID)
        {
            //CompanyManager.getSingleton().init();
            string parentPath = CompanyManager.getSingleton().GetCompanyFullName(CompanyID);
            parentPath = parentPath.Replace("/", "\\");
            return HttpRuntime.AppDomainAppPath + parentPath + "组织结构图.jpg";
        }

    }
}
