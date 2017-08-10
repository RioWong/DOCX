using CoreLib;
using CreateWord.DB;
using CreateWord.log;
using RealEstate.DAL.SQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace CreateWord.model
{
    //房地信息统计模块里面需要用到的对象

    #region  “县级”房产信息统计模块的表格数据类
    public class FDXXtbl_country
    {
        #region 成员变量
        //地块ID
        public int ZDXX_ID { get; set; }
        //地块名称
        public string ZDXX_MC { get; set; }
        //地块占地面积
        public decimal ZDXX_ZDMJ { get; set; }
        //建筑ID
        public int FCXX_ID { get; set; }
        //建筑名称
        public string FCXX_JZMC { get; set; }
        //建筑层数地上
        public decimal FCXX_DSCS { get; set; }
        //建筑层数地下
        public decimal FCXX_DXCS { get; set; }
        //建筑结构
        public string FCXX_JZJG { get; set; }
        //建筑年代
        public decimal FCXX_JSND { get; set; }
        //建筑面积总体
        public decimal ZMJ { get; set; }
        //建筑面积地上
        public decimal FCXX_DSMJ { get; set; }
        //建筑面积地下
        public decimal FCXX_DXMJ { get; set; }
        //具体使用功能
        public string FCZK_SYGN { get; set; }
        //具体使用部门
        public string FCZK_SYBM { get; set; }
        //备注
        public string FCXX_BZ { get; set; }
        #endregion

        #region 相关操作

        /// <summary>
        /// 从数据库中获取一行行表格数据，并存入FDXXmodel中。（县级公司）
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static List<FDXXtbl_country> GetInfo(int id)
        {
            List<FDXXtbl_country> lstFDXX = new List<FDXXtbl_country>();
            FDXXtbl_country fm;
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                
                string sql = CompilationDocumentSQL.GetFDXXcountrySQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        fm = new FDXXtbl_country();
                        fm.ZDXX_ID = int.Parse(((decimal)myReader["ZDXX_ID"]).ToString()); //(int)myReader["ZDXX_ID"]; 
                        fm.ZDXX_MC = (string)myReader["ZDXX_MC"];
                        fm.ZDXX_ZDMJ = Decimal.Parse((((decimal)myReader["ZDXX_ZDMJ"]) * 2000 / 3).ToString("0"));    //把单位统一为平方米
                        fm.FCXX_ID = int.Parse(((decimal)myReader["FCXX_ID"]).ToString()); //(int)myReader["FCXX_ID"];
                        fm.FCXX_JZMC = (string)myReader["FCXX_JZMC"];
                        fm.FCXX_DSCS = (decimal)myReader["FCXX_DSCS"];
                        fm.FCXX_DXCS = (decimal)myReader["FCXX_DXCS"];
                        fm.FCXX_JZJG = (string)myReader["FCXX_JZJG"];
                        fm.FCXX_JSND = (decimal)myReader["FCZK_JSND"];
                        fm.ZMJ = Math.Round((decimal)myReader["ZMJ"], 2);
                        fm.FCXX_DSMJ = Math.Round((decimal)myReader["FCXX_DSMJ"], 2);
                        fm.FCXX_DXMJ = Math.Round((decimal)myReader["FCXX_DXMJ"], 2);
                        fm.FCZK_SYGN = (string)myReader["FCZK_SYGN"];
                        fm.FCZK_SYBM = (string)myReader["FCZK_SYBM"];
                        if (!DBNull.Value.Equals(myReader["FCXX_BZ"])) fm.FCXX_BZ = (string)myReader["FCXX_BZ"];//判断FCXX_BZ是否为空值（DBNull）
                        else fm.FCXX_BZ = "";

                        lstFDXX.Add(fm);
                    }
                }
                
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_country), ex);
            }
            finally
            {
                mycon.Close();
            }
            return lstFDXX;
        }

        /// <summary>
        /// 读取合计信息。（县级公司）
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static FDXXtbl_country GetTotalInfo(int id)
        {
            FDXXtbl_country fm= new FDXXtbl_country();
            IDbConnection mycon = DBhelper.OpenConnection();
            try
            {                
                string sql = CompilationDocumentSQL.GetFDXXcountrytotalSQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using (IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    myReader.Read();//把得到的值赋给fm对象

                    fm.ZDXX_ZDMJ = Math.Round((decimal)myReader["ZDMJ"], 0);
                    fm.ZMJ = Math.Round((decimal)myReader["FCMJ"], 0);
                    fm.FCXX_DSMJ = Math.Round((decimal)myReader["DSMJ"], 0);
                    fm.FCXX_DXMJ = Math.Round((decimal)myReader["DXMJ"], 0);
                }               

            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_country), ex);
            }
            finally
            {
                mycon.Close();
            }
            return fm;
        }

        /// <summary>
        /// 得到ZDXX_MC这个字段名称重复的个数，用于合并单元格。
        /// </summary>
        /// <param name="lstFDXX"></param>
        /// <returns>返回无重复的宗地名称和宗地出现的次数</returns>
        public static List<Parcelmodel> Parcels(List<FDXXtbl_country> lstFDXX)
        {
            int flag;
            int num = 0;
            string temp = "";
            Parcelmodel pm;
            List<Parcelmodel> lstParcel = new List<Parcelmodel>();

            try
            {
                foreach (FDXXtbl_country fm in lstFDXX)
                {
                    num = 0;
                    flag = 1;
                    temp = fm.ZDXX_MC;
                    foreach (Parcelmodel pm1 in lstParcel)
                    {
                        if (pm1.name == temp) { flag = 0; break; }//说明p[]里面已经存储了该宗地信息，所以把flag值设置为0。
                    }
                    if (flag == 1)
                    {
                        foreach (FDXXtbl_country fm1 in lstFDXX)
                        {
                            if (fm1.ZDXX_MC == temp) num++;
                        }
                        pm = new Parcelmodel();
                        pm.id = fm.ZDXX_ID;
                        pm.name = temp;
                        pm.num = num;
                        lstParcel.Add(pm);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_country), ex);
            }

            return lstParcel;
        }

        public static List<Buildingmodel> Buildings(List<FDXXtbl_country> lstFDXX)
        {
            int flag;
            int num = 0;
            //string temp = "";
            Buildingmodel pm;
            List<Buildingmodel> lstBuilding = new List<Buildingmodel>();

            try
            {
                foreach (FDXXtbl_country fm in lstFDXX)
                {
                    num = 0;
                    flag = 1;
                    //temp = fm.FCXX_JZMC;
                    int buildingId = fm.FCXX_ID;
                    foreach (Buildingmodel pm1 in lstBuilding)
                    {
                        //if (pm1.name == temp) { flag = 0; break; }//说明p[]里面已经存储了该宗地信息，所以把flag值设置为0。
                        if (pm1.id == buildingId) { flag = 0; break; }//说明p[]里面已经存储了该宗地信息，所以把flag值设置为0。
                    }
                    if (flag == 1)
                    {
                        foreach (FDXXtbl_country fm1 in lstFDXX)
                        {
                            //if (fm1.FCXX_JZMC == temp) num++;
                            if (fm1.FCXX_ID == buildingId) num++;
                        }
                        pm = new Buildingmodel();
                        pm.id = fm.FCXX_ID;
                        pm.name = fm.FCXX_JZMC;// temp;
                        pm.num = num;
                        lstBuilding.Add(pm);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_country), ex);
            }

            return lstBuilding;
        }

        #endregion
    }
    #endregion

    #region “市级”房产信息统计模块的表格数据类
    public class FDXXtbl_city
    {
        #region 成员变量
        //单位名称
        public string DWMC { get; set; }
        //宗地数
        public decimal ZDS { get; set; }
        //宗地面积
        public decimal ZDMJ { get; set; }
        //房产数
        public decimal FCS { get; set; }
        //房产面积
        public decimal FCMJ { get; set; }
        #endregion

        #region 相关操作

        /// <summary>
        /// 从数据库中获取一行行表格数据，并存入FDXXmodel中。（市级公司）
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static List<FDXXtbl_city> GetInfo(int id)
        {
            List<FDXXtbl_city> lstFDXX = new List<FDXXtbl_city>();
            FDXXtbl_city fm;
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                
                string sql = CompilationDocumentSQL.GetFDXXcitySQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        fm = new FDXXtbl_city();
                        fm.DWMC = (string)myReader["DWMC"];
                        fm.ZDS = (decimal)myReader["ZDS"];
                        fm.ZDMJ = (decimal)myReader["ZDMJ"];
                        fm.FCS = (decimal)myReader["FCS"];
                        fm.FCMJ = (decimal)myReader["FCMJ"];

                        lstFDXX.Add(fm);
                    }
                }                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_city), ex);
            }
            finally
            {
                mycon.Close();
            }
            return lstFDXX;
        }
        #endregion
    }
    #endregion

    #region “省级”房产信息统计模块的表格数据类
    /// <summary>
    /// 总汇表
    /// </summary>
    public class FDXXtbl_province_summary
    {
        #region 成员变量
        //单位名称
        public string DWMC { get; set; }
        //地块占地面积
        public decimal ZDMJ { get; set; }
        //建筑面积（总体）
        public decimal FCMJ { get; set; }
        //建筑面积（地上）
        public decimal DSMJ { get; set; }
        //建筑面积（地下）
        public decimal DXMJ { get; set; }
        #endregion

        #region 相关操作
        /// <summary>
        /// 从数据库中获取一行行表格数据，并存入FDXXmodel中。（省公司公司）
        /// </summary>
        /// <returns></returns>
        public static List<FDXXtbl_province_summary> GetInfo()
        {
            List<FDXXtbl_province_summary> lstFDXX = new List<FDXXtbl_province_summary>();
            FDXXtbl_province_summary fm;
            IDbConnection mycon = DBhelper.OpenConnection();          

            try
            {
                
                string sql = CompilationDocumentSQL.GetFDXXprovincesummarySQL();
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        fm = new FDXXtbl_province_summary();
                        fm.DWMC = (string)myReader["DWMC"];
                        fm.ZDMJ = Decimal.Parse((((decimal)myReader["ZDMJ"]) * (decimal)0.0015).ToString("0"));    //把单位统一为亩
                        fm.FCMJ = (decimal)myReader["FCMJ"];
                        fm.DSMJ = (decimal)myReader["DSMJ"];
                        fm.DXMJ = (decimal)myReader["DXMJ"];
                        lstFDXX.Add(fm);
                    }
                }
                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_province_summary), ex);
            }
            finally
            {
                mycon.Close();
            }
            return lstFDXX;
        }
        #endregion
    }

    /// <summary>
    /// 分析表
    /// </summary>
    public class FDXXtbl_province_analysis
    {
        #region 成员变量
        //单位
        public string DW { get; set; }
        //统计项目名称
        public string TJXM { get; set; }
        //建筑面积
        public double FCMJ { get; set; }
        //用于排序
        public int order;
        #endregion

        #region 相关操作
        public static List<FDXXtbl_province_analysis> GetInfo()
        {
            List<FDXXtbl_province_analysis> lstFDXX = new List<FDXXtbl_province_analysis>();
            IDbConnection mycon = DBhelper.OpenConnection();
            try
            {               
                
                string sql = CompilationDocumentSQL.GetFDXXprovinceanalysisSQL();
                DataTable t = DBhelper.GetDatatable(sql, mycon);

                //DataTable t = DataHelper.Get_datatable(mycon, sql);
                /*------------------------------排序------------------------------*/
                //将datatable存入自定义数组
                for (int i = 0; i < t.Rows.Count; i++)
                {
                    FDXXtbl_province_analysis fa = new FDXXtbl_province_analysis();
                    fa.DW = Convert.ToString(t.Rows[i][0]);
                    fa.TJXM = Convert.ToString(t.Rows[i][1]);
                    fa.FCMJ = Convert.ToDouble(t.Rows[i][2]);
                    switch (fa.TJXM)
                    {
                        case "合计": fa.order = 1; break;
                        case "本部": fa.order = 2; break;
                        case "供电公司": fa.order = 3; break;
                        case "培训单位": fa.order = 4; break;
                        case "直属单位": fa.order = 5; break;
                        case "调度生产管理用房": fa.order = 6; break;
                        case "营销服务用房": fa.order = 7; break;
                        case "运维检修用房": fa.order = 8; break;
                        case "物资仓储用房": fa.order = 9; break;
                        case "教育培训用房": fa.order = 10; break;
                        case "科研实验用房": fa.order = 11; break;
                        case "其它用房": fa.order = 12; break;
                        case "10年以内（含10年）": fa.order = 13; break;
                        case "10-20年（含20年）": fa.order = 14; break;
                        case "20-30年（含30年）": fa.order = 15; break;
                        case "30年以上": fa.order = 16; break;
                        case "危房": fa.order = 17; break;
                        case "正常": fa.order = 18; break;
                        case "规划拆除": fa.order = 19; break;
                        case "尚未规划拆除": fa.order = 20; break;
                        case "已办理土地证": fa.order = 21; break;
                        case "未办理土地证": fa.order = 22; break;
                        case "已办理房产证": fa.order = 23; break;
                        case "未办理房产证": fa.order = 24; break;
                        default:
                            break;
                    }
                    lstFDXX.Add(fa);
                }
                //对List中的表数据进行排序
                lstFDXX = lstFDXX.OrderBy(s => s.order).ToList<FDXXtbl_province_analysis>();
            }
            catch(Exception ex)
            {
                LogHelper.WriteLog(typeof(FDXXtbl_province_analysis), ex);
            }
            finally
            {
                mycon.Close();
            }
            return lstFDXX;
        }
        #endregion

    }


    #endregion

    #region “县级”房地信息统计模块的文字描述类
    /// <summary>
    /// 第一句话
    /// </summary>
    public class countryFDXXsentence1
    {
        #region 成员变量
        //各类住房栋数
        public decimal count { get; set; }
        //占地总面积
        public decimal ZDZMJ { get; set; }
        //总建筑面积
        public decimal ZJZMJ { get; set; }
        #endregion

        /// <summary>
        /// 从数据库中得到各类住房栋数、占地总面积、总建筑面积并存入countryFDXXsentence1对象中。
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static countryFDXXsentence1 GetInfo(int id)
        {
            countryFDXXsentence1 fs = new countryFDXXsentence1();
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {

                
                string sql = CompilationDocumentSQL.GetcountryFDXXsentence1SQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    myReader.Read();
                    fs.count = (decimal)myReader["FCS"];
                    fs.ZDZMJ = (decimal)myReader["ZDXX_ZDMJ"];
                    fs.ZJZMJ = (decimal)myReader["B_FCMJ"];
                }
                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(countryFDXXsentence1), ex);
            }
            finally
            {
                mycon.Close();
            }
            return fs;

        }
    }

    /// <summary>
    /// 第二句话
    /// </summary>
    public class countryFDXXsentence2
    {
        #region 成员变量
        //各类用房名称
        public string GNGL { get; set; }
        //各类用房面积
        public decimal GNGL_MJ { get; set; }
        #endregion

        /// <summary>
        /// 从数据库中获取各类用房名称和面积存入FDXXsentence2对象中。
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static List<countryFDXXsentence2> GetInfo(int id)
        {
            List<countryFDXXsentence2> lstFS = new List<countryFDXXsentence2>();
            countryFDXXsentence2 fs;
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                
                string sql = CompilationDocumentSQL.GetcountryFDXXsentence2SQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        fs = new countryFDXXsentence2();
                        fs.GNGL = (string)myReader["FCZK_GNGL"];
                        fs.GNGL_MJ = (decimal)myReader["SUM(FCMJ)"];

                        lstFS.Add(fs);
                    }
                }
                
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(countryFDXXsentence2), ex);
            }
            finally
            {
                mycon.Close();
            }
            return lstFS;
        }


    }

    /// <summary>
    /// 第三句话
    /// </summary>
    public class countryFDXXsentence3
    {
        #region 成员变量
        //建成投运10年内的房屋面积
        public decimal FWMJ_10 { get; set; }
        //建成投运10-20年内的房屋面积
        public decimal FWMJ_1020 { get; set; }
        //建成投运20-30年内的房屋面积
        public decimal FWMJ_2030 { get; set; }
        //建成投运30年以上的房屋面积
        public decimal FWMJ_30 { get; set; }
        #endregion

        /// <summary>
        /// 从数据库中获取建成投运10年内的房屋面积、建成投运10-20年内的房屋面积和建成投运30年以上的房屋面积并存入FDXXsentence3对象中。
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static countryFDXXsentence3 GetInfo(int id)
        {
            countryFDXXsentence3 fs = new countryFDXXsentence3();
            string temp = "";
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                
                string sql = CompilationDocumentSQL.GetcountryFDXXsentence3SQL(id);
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader  = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        temp = (string)myReader["TJXM"];
                        if (temp == "10年以内（含10年）") fs.FWMJ_10 = (decimal)myReader["FCMJ"];
                        else if (temp == "10-20年（含20年）") fs.FWMJ_1020 = (decimal)myReader["FCMJ"];
                        else if (temp == "20-30年（含30年）") fs.FWMJ_2030 = (decimal)myReader["FCMJ"];
                        else fs.FWMJ_30 = (decimal)myReader["FCMJ"];
                    }
                }
                
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(countryFDXXsentence3), ex);
            }
            finally
            {
                mycon.Close();
            }
            return fs;
        }
    }
    #endregion

    #region “省级”房地信息统计模块的文字描述类
    /// <summary>
    /// 第一句话
    /// </summary>
    public class provinceFDXXsentence1
    {
        #region 成员变量
        //对应书签名：土地总面积
        public decimal ZDMJ { get; set; }
        //对应书签名：用房总面积
        public decimal FCMJ { get; set; }
        //对应书签名：地上面积
        public decimal DSMJ { get; set; }
        //对应书签名：地下面积
        public decimal DXMJ { get; set; }
        #endregion

        /// <summary>
        /// 从数据库中得到土地总面积、用房总面积、地上面积、地下面积并存入provinceFDXXsentence1对象中。
        /// </summary>
        /// <returns></returns>
        public static provinceFDXXsentence1 GetInfo()
        {
            provinceFDXXsentence1 fs = new provinceFDXXsentence1();
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                
                string sql = CompilationDocumentSQL.GetprovinceFDXXsentence1SQL();
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    myReader.Read();
                    fs.ZDMJ = Decimal.Parse((((decimal)myReader["ZDMJ"]) * (decimal)0.0015).ToString("0"));    //把单位统一为亩
                    fs.FCMJ = (decimal)myReader["FCMJ"];
                    fs.DSMJ = (decimal)myReader["DSMJ"];
                    fs.DXMJ = (decimal)myReader["DXMJ"];
                }
                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(provinceFDXXsentence1), ex);
            }
            finally
            {
                mycon.Close();
            }
            return fs;
        }
    }

    /// <summary>
    /// 第二句话
    /// </summary>
    public class provinceFDXXsentence2
    {
        #region 成员变量
        //对应书签名：危房面积
        public decimal WFZMJ { get; set; }
        //对应书签名：危房面积比
        public string WFZMJB { get; set; }
        //对应书签名：规划拆除用房面积
        public decimal CCZMJ { get; set; }
        //对应书签名：规划拆除用房面积比
        public string CCZMJB { get; set; }
        //对应书签名：未办权证用房面积
        public decimal NQZZMJ { get; set; }
        //对应书签名：未办权证用房面积比
        public string NQZZMJB { get; set; }
        //对应书签名：二零一三年以前的房屋面积
        public decimal BF2013ZMJ { get; set; }
        //对应书签名：二零一三年以前的房屋面积比
        public string BF2013ZMJB { get; set; }
        #endregion

        public static provinceFDXXsentence2 GetInfo()
        {
            provinceFDXXsentence1 fs1 = new provinceFDXXsentence1();
            provinceFDXXsentence2 fs = new provinceFDXXsentence2();
            string temp;
            decimal ZMJ;
            IDbConnection mycon = DBhelper.OpenConnection();
            
            try
            {
                fs1 = provinceFDXXsentence1.GetInfo();
                ZMJ = fs1.FCMJ;
                
                string sql = CompilationDocumentSQL.GetprovinceFDXXsentence2SQL();
                IDbCommand mycom = DBFactory.getSingleton().getCommand(sql, mycon);
                using( IDataReader myReader  = mycom.ExecuteReader())//执行command并得到相应的DataReader
                {
                    while (myReader.Read())//把得到的值赋给fm对象
                    {
                        temp = (string)myReader["MC"];
                        if (temp == "危房")
                        {
                            fs.WFZMJ = (decimal)myReader["ZMJ"];
                            fs.WFZMJB = (Math.Round(fs.WFZMJ / ZMJ * 100, 2)).ToString() + "%";
                        }
                        else if (temp == "拆除")
                        {
                            fs.CCZMJ = (decimal)myReader["ZMJ"];
                            fs.CCZMJB = (Math.Round(fs.CCZMJ / ZMJ * 100, 2)).ToString() + "%";
                        }
                        else if (temp == "未办权证")
                        {
                            fs.NQZZMJ = (decimal)myReader["ZMJ"];
                            fs.NQZZMJB = (Math.Round(fs.NQZZMJ / ZMJ * 100, 2)).ToString() + "%";
                        }
                        else if (temp == "2013年以前")
                        {
                            fs.BF2013ZMJ = (decimal)myReader["ZMJ"];
                            fs.BF2013ZMJB = (Math.Round(fs.BF2013ZMJ / ZMJ * 100, 2)).ToString() + "%";
                        }
                    }
                }
                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(provinceFDXXsentence2), ex);
            }
            finally
            {
                mycon.Close();
            }
            return fs;
        }

    }
    #endregion

}
