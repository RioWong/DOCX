using CoreLib;
using CreateWord.model;
using RealEstate.DAL;
using RealEstate.Logic;
using RealEstate.Model;
using System.Collections.Generic;
using System.Data;

namespace CreateWord.DB
{
    public class DBhelper
    {
        ///// <summary>
        ///// 数据库连接字符串
        ///// </summary>
        //public static string ConnectionString
        //{
        //    get { return ConfigurationManager.ConnectionStrings["JSDL"].ConnectionString; }
        //}

        /// <summary>
        /// 打开数据库
        /// </summary>
        /// <returns></returns>
        public static IDbConnection OpenConnection()
        {
            return CommonDAL.OpenConnection();
        }

        /// <summary>
        /// 根据公司ID获取公司名称
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static string GetCompanyName(int id)
        {            
            return CompanyManager.getSingleton().GetCompanyName(id);
        }
        /// <summary>
        /// 根据公司ID获取其子公司的ID和名称
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static List<Companymodel> GetChildcompany(int id)
        {
            List<Companymodel> lstCM = new List<Companymodel>();
            IList<CompanyBase> lst = CompanyManager.getSingleton().GetCompanyList(id, false, false);
            foreach( CompanyBase c in lst )
            {
                lstCM.Add( new Companymodel(){ name = c.Name, ID = c.ID, property = c.Property });
            }
            return lstCM;
        }

        public static DataTable GetDatatable(string sql, IDbConnection Con)
        {
            IDbDataAdapter da = DBFactory.getSingleton().getDataAdapter(sql, Con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            DataTable dt = ds.Tables[0];
            return dt;
        }
    }
}
