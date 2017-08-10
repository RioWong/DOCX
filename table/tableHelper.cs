using CreateWord.log;
using CreateWord.model;
using Novacode;
using RealEstate.Logic;
using RealEstate.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.table
{

    /// <summary>
    /// 表相关的操作
    /// </summary>
    class tableHelper
    {

        /// <summary>
        /// 创建"县级公司"房地信息汇总表模板（前两行，字段；第三行，总计）
        /// </summary>
        /// <returns></returns>
        public static Table Template_country(DocX document)
        {
            Table t = null;
            try
            {
                t = document.AddTable(3, 14);

                #region "前三行"
                t.Rows[0].Cells[0].Paragraphs.First().Append("序号").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs.First().Append("地块名称").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[2].Paragraphs.First().Append("地块占地面积（m2）").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs.First().Append("建筑名称").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[4].Paragraphs.First().Append("建筑层数").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[6].Paragraphs.First().Append("建筑结构").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[7].Paragraphs.First().Append("建筑年代").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[8].Paragraphs.First().Append("建筑面积（m2）").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[11].Paragraphs.First().Append("具体使用功能").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[12].Paragraphs.First().Append("使用部门").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[13].Paragraphs.First().Append("备注").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[4].Paragraphs.First().Append("地上").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[5].Paragraphs.First().Append("地下").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[8].Paragraphs.First().Append("总体").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[9].Paragraphs.First().Append("地上").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[10].Paragraphs.First().Append("地下").Bold().Alignment = Alignment.center;
                t.Rows[2].Cells[0].Paragraphs.First().Append("合计").Bold().Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;//垂直居中
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[6].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[7].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[8].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[11].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[12].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[13].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[5].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[8].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[9].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[10].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[2].Cells[0].VerticalAlignment = VerticalAlignment.Center;

                //单元格合并操作（先竖向合并，再横向合并，以免报错，因为横向合并会改变列数）
                t.MergeCellsInColumn(0, 0, 1);
                t.MergeCellsInColumn(1, 0, 1);
                t.MergeCellsInColumn(2, 0, 1);
                t.MergeCellsInColumn(3, 0, 1);
                t.MergeCellsInColumn(6, 0, 1);
                t.MergeCellsInColumn(7, 0, 1);
                t.MergeCellsInColumn(11, 0, 1);
                t.MergeCellsInColumn(12, 0, 1);
                t.MergeCellsInColumn(13, 0, 1);
                t.Rows[0].MergeCells(4, 5);
                t.Rows[0].MergeCells(7, 9);
                #endregion
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }


        /// <summary>
        /// 把数据一行行插入（县级公司）
        /// </summary>
        /// <param name="t"></param>
        /// <param name="lsmFM"></param>
        /// <returns></returns>
        public static Table inserttable_country(Table t, List<FDXXtbl_country> lsmFM,FDXXtbl_country FM)
        {
            int temp = 2;
            try
            {
                t.Rows[temp].Cells[2].Paragraphs.First().Append("" + FM.ZDXX_ZDMJ).Alignment = Alignment.center;
                t.Rows[temp].Cells[8].Paragraphs.First().Append("" + FM.ZMJ).Alignment = Alignment.center;
                t.Rows[temp].Cells[9].Paragraphs.First().Append("" + FM.FCXX_DSMJ).Alignment = Alignment.center;
                t.Rows[temp].Cells[10].Paragraphs.First().Append("" + FM.FCXX_DXMJ).Alignment = Alignment.center;

                t.Rows[temp].Cells[2].VerticalAlignment = VerticalAlignment.Center;//垂直居中
                t.Rows[temp].Cells[8].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[temp].Cells[9].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[temp].Cells[10].VerticalAlignment = VerticalAlignment.Center;
                temp++;
                foreach (FDXXtbl_country fm in lsmFM)
                {
                    t.InsertRow();
                    t.Rows[temp].Cells[1].Paragraphs.First().Append(fm.ZDXX_MC).Alignment = Alignment.center;
                    t.Rows[temp].Cells[2].Paragraphs.First().Append("" + fm.ZDXX_ZDMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[3].Paragraphs.First().Append(fm.FCXX_JZMC).Alignment = Alignment.center;
                    t.Rows[temp].Cells[4].Paragraphs.First().Append("" + fm.FCXX_DSCS).Alignment = Alignment.center;
                    t.Rows[temp].Cells[5].Paragraphs.First().Append("" + fm.FCXX_DXCS).Alignment = Alignment.center;
                    t.Rows[temp].Cells[6].Paragraphs.First().Append(fm.FCXX_JZJG).Alignment = Alignment.center;
                    t.Rows[temp].Cells[7].Paragraphs.First().Append("" + fm.FCXX_JSND).Alignment = Alignment.center;
                    t.Rows[temp].Cells[8].Paragraphs.First().Append("" + fm.ZMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[9].Paragraphs.First().Append("" + fm.FCXX_DSMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[10].Paragraphs.First().Append("" + fm.FCXX_DXMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[11].Paragraphs.First().Append(fm.FCZK_SYGN).Alignment = Alignment.center;
                    t.Rows[temp].Cells[12].Paragraphs.First().Append(fm.FCZK_SYBM).Alignment = Alignment.center;
                    t.Rows[temp].Cells[13].Paragraphs.First().Append(fm.FCXX_BZ).Alignment = Alignment.center;

                    t.Rows[temp].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[5].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[6].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[7].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[8].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[9].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[10].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[11].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[12].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[13].VerticalAlignment = VerticalAlignment.Center;

                    temp++;
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 创建"市级公司"房地信息汇总表模板（两行，字段）
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Table Template_city(DocX document)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 5);

                #region "前两行"
                t.Rows[0].Cells[0].Paragraphs.First().Append("单位名称").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs.First().Append("地产信息").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs.First().Append("房产信息").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[1].Paragraphs.First().Append("宗地数（个）").Bold().Color(Color.Red).Alignment = Alignment.center;
                t.Rows[1].Cells[2].Paragraphs.First().Append("宗地面积（m2）").Bold().Color(Color.Red).Alignment = Alignment.center;
                t.Rows[1].Cells[3].Paragraphs.First().Append("房产数（个）").Bold().Color(Color.Red).Alignment = Alignment.center;
                t.Rows[1].Cells[4].Paragraphs.First().Append("房产面积（m2）").Bold().Color(Color.Red).Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[4].VerticalAlignment = VerticalAlignment.Center;



                //单元格合并操作（先竖向合并，再横向合并，以免报错，因为横向合并会改变列数）
                t.MergeCellsInColumn(0, 0, 1);
                t.Rows[0].MergeCells(1, 2);
                t.Rows[0].MergeCells(2, 3);
                #endregion
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }




        /// <summary>
        /// 把数据一行行插入（市级公司）
        /// </summary>
        /// <param name="t"></param>
        /// <param name="lsmFM"></param>
        /// <returns></returns>
        public static Table inserttable_city(Table t, List<FDXXtbl_city> lsmFM)
        {
            int temp = 2;
            try
            {
                foreach (FDXXtbl_city fm in lsmFM)
                {
                    t.InsertRow();
                    t.Rows[temp].Cells[0].Paragraphs.First().Append(fm.DWMC).Alignment = Alignment.center;
                    t.Rows[temp].Cells[1].Paragraphs.First().Append("" + fm.ZDS).Alignment = Alignment.center;
                    t.Rows[temp].Cells[2].Paragraphs.First().Append("" + fm.ZDMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[3].Paragraphs.First().Append("" + fm.FCS).Alignment = Alignment.center;
                    t.Rows[temp].Cells[4].Paragraphs.First().Append("" + fm.FCMJ).Alignment = Alignment.center;

                    

                    t.Rows[temp].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                    temp++;
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 创建"省公司"房地信息汇总表模板（两行，字段）
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Table Template_province_summary(DocX document)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 6);

                #region "前两行"
                t.Rows[0].Cells[0].Paragraphs.First().Append("序号").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs.First().Append("单位名称").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[2].Paragraphs.First().Append("地块占地面积（亩）").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs.First().Append("建筑面积（m2）").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[3].Paragraphs.First().Append("总体").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[4].Paragraphs.First().Append("地上").Bold().Alignment = Alignment.center;
                t.Rows[1].Cells[5].Paragraphs.First().Append("地下").Bold().Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[5].VerticalAlignment = VerticalAlignment.Center;


                //单元格合并操作（先竖向合并，再横向合并，以免报错，因为横向合并会改变列数）
                t.MergeCellsInColumn(0, 0, 1);
                t.MergeCellsInColumn(1, 0, 1);
                t.MergeCellsInColumn(2, 0, 1);
                t.Rows[0].MergeCells(3, 5);
                #endregion
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 把数据一行行插入（省公司）
        /// </summary>
        /// <param name="t"></param>
        /// <param name="lsmFM"></param>
        /// <returns></returns>
        public static Table inserttable_province_summary(Table t, List<FDXXtbl_province_summary> lsmFM)
        {
            int temp = 2;
            int num = 0;
            bool flag = true;
            try
            {
                foreach (FDXXtbl_province_summary fm in lsmFM)
                {
                    t.InsertRow();
                    if (num != 0)
                    {
                        if (num == 1 && flag == true)
                        {
                            num--; flag = false;
                        }

                        t.Rows[temp].Cells[0].Paragraphs.First().Append("" + ++num).Alignment = Alignment.center;

                        t.Rows[temp].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    }
                    else num++; 
                    t.Rows[temp].Cells[1].Paragraphs.First().Append(fm.DWMC).Alignment = Alignment.center;
                    t.Rows[temp].Cells[2].Paragraphs.First().Append("" + fm.ZDMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[3].Paragraphs.First().Append("" + fm.FCMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[4].Paragraphs.First().Append("" + fm.DSMJ).Alignment = Alignment.center;
                    t.Rows[temp].Cells[5].Paragraphs.First().Append("" + fm.DXMJ).Alignment = Alignment.center;

                    t.Rows[temp].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[4].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[5].VerticalAlignment = VerticalAlignment.Center;

                    temp++;
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 创建“省公司”的企业用房分析表（前两行）
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Table Template_province_analysis(DocX document)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 4);

                #region "前两行"
                t.Rows[0].Cells[0].Paragraphs.First().Append("序号").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs.First().Append("单位").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[2].Paragraphs.First().Append("统计项目名称").Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs.First().Append("建筑面积（m2）").Bold().Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[3].VerticalAlignment = VerticalAlignment.Center;

                #endregion


            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 把数据一行行插入（省公司）
        /// </summary>
        /// <param name="t"></param>
        /// <param name="lsmFM"></param>
        /// <returns></returns>
        public static Table inserttable_province_analysis(Table t, List<FDXXtbl_province_analysis> lsmFM)
        {
            int temp = 1;
            int num = 0;
            try
            {
                foreach (FDXXtbl_province_analysis fm in lsmFM)
                {
                    t.InsertRow();
                    if (num != 0)
                    {
                        num++;
                        t.Rows[temp].Cells[0].Paragraphs.First().Append("" + num).Alignment = Alignment.center;

                        t.Rows[temp].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    }
                    t.Rows[temp].Cells[1].Paragraphs.First().Append(fm.DW).Alignment = Alignment.center;
                    t.Rows[temp].Cells[2].Paragraphs.First().Append(fm.TJXM).Alignment = Alignment.center;
                    t.Rows[temp].Cells[3].Paragraphs.First().Append("" + fm.FCMJ).Alignment = Alignment.center;

                    t.Rows[temp].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    t.Rows[temp].Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    temp++;
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 合并重复单位名称（省公司的用房分析表）
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public static Table combineCells(Table t)
        {
            string tempDW = "";
            int loc = 0;//需要合并的单元格第一个的位置
            int num = 1;//统计需要合并的单元格数量
            int number = 1;//表中第一列的序号

            for (int i = 1; i < t.RowCount; i++)
            {
                if (tempDW != t.Rows[i].Cells[1].Paragraphs.First().Text)
                {
                    if (tempDW != "")
                    {
                        t.MergeCellsInColumn(0, loc, loc + num - 1);//合并序号
                        t.MergeCellsInColumn(1, loc, loc + num - 1);//合并单位
                        t.Rows[loc].Cells[0].Paragraphs.First().Append("" + number++).Alignment = Alignment.center;//添加序号

                        t.Rows[loc].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    }
                    tempDW = t.Rows[i].Cells[1].Paragraphs.First().Text;
                    num = 1;
                    loc = i;
                }
                else num++;
            }
            return t;
        }

        /// <summary>
        /// 合并重复的地块名称（县级公司的汇总表）
        /// </summary>
        /// <param name="t"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static Table combineCells(Table t, List<Parcelmodel> p, List<Buildingmodel> b)
        {
            int startCell = 3;
            int number = 1;//表中第一列的序号

            try
            {
                foreach (Parcelmodel pm in p)
                {
                    if (pm.num != 1)
                    {
                        t.MergeCellsInColumn(0, startCell, startCell + pm.num - 1); //合并序号                    
                        t.MergeCellsInColumn(1, startCell, startCell + pm.num - 1); //合并地块名称
                        t.MergeCellsInColumn(2, startCell, startCell + pm.num - 1); //合并地块占地面积
                    }
                    t.Rows[startCell].Cells[0].Paragraphs.First().Append("" + number++).Alignment = Alignment.center;    //添加序号
                    t.Rows[startCell].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    startCell = startCell + pm.num;
                }

                startCell = 3;
                foreach(Buildingmodel bm in b)
                {
                    if(bm.num != 1)
                    {
                        t.MergeCellsInColumn(3, startCell, startCell + bm.num - 1); //合并建筑名称                  
                        t.MergeCellsInColumn(4, startCell, startCell + bm.num - 1); //合并地上层数
                        t.MergeCellsInColumn(5, startCell, startCell + bm.num - 1); //合并地下层数
                        t.MergeCellsInColumn(6, startCell, startCell + bm.num - 1); //合并建筑结构                 
                        t.MergeCellsInColumn(7, startCell, startCell + bm.num - 1); //合并建筑年代
                        t.MergeCellsInColumn(8, startCell, startCell + bm.num - 1); //合并总体面积
                        t.MergeCellsInColumn(9, startCell, startCell + bm.num - 1); //合并地上面积                
                        t.MergeCellsInColumn(10, startCell, startCell + bm.num - 1); //合并地下面积
                        
                    }
                    startCell = startCell + bm.num;
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }

            return t;
        }

        /// <summary>
        /// 创建表，用来展示平面图（Plan）的表。
        /// </summary>
        /// <returns></returns>
        public static Table PlanTable(DocX document, int CompanyID, int parcelId, int buildingId, List<string> lstStr)
        {
            Table t = null;
            int i = 0;

            try
            {
                if (lstStr.Count == 0)
                    t = document.AddTable(1, 1);
                else
                {
                    t = document.AddTable(lstStr.Count, 1);
                    foreach (string s in lstStr)
                    {
                        Picture p = picture.picHelper.getPic(document, PathManager.getSingleton().GetPlanpicPath( buildingId, s, false), 287, 670);
                        if (p == null) continue;

                        t.Rows[i].Cells[0].Paragraphs.First().AppendPicture(p).AppendLine(s).FontSize(16).Bold().Alignment = Alignment.center;
                        t.Rows[i].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                        t.Rows[i++].Cells[0].Width = 23.27;
                    }
                }                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }

            return t;
        }

        /// <summary>
        /// 创建市公司的组织结构表
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childCompanies"></param>
        /// <param name="picPath"></param>
        /// <returns></returns>
        public static Table organizationTable(DocX document, List<Companymodel> childCompanies, string picPath)
        {
            if (childCompanies.Count == 0) return null;

            Table t = document.AddTable(childCompanies.Count + 1, 2);
            

            int i = 0;
            foreach (Companymodel cm in childCompanies)
            {
                if (i % 2 == 0)
                    t.Rows[i].Cells[0].FillColor = Color.Blue;
                else t.Rows[i].Cells[0].FillColor = Color.LightBlue;
                t.Rows[i].Cells[0].Paragraphs.First().Append(cm.name).FontSize(15).Bold().Alignment = Alignment.center;
                t.Rows[i++].Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            t.MergeCellsInColumn(1, 0, childCompanies.Count);
            Picture p = picture.picHelper.getPic(document, picPath, 728, 525);
            if (p != null)
            {
                t.Rows[0].Cells[1].Paragraphs.First().AppendPicture(p).Alignment = Alignment.center;//插入组织结构图
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
            }
            return t;

        }

        /// <summary>
        /// 生成用于放置县公司位置分布图的表格
        /// </summary>
        /// <param name="document"></param>
        /// <param name="p1"></param>
        /// <returns></returns>
        public static Table locpicTable(DocX document, Picture p1)
        {
            Table t = null;
            try
            {
                t = document.AddTable(1, 1);
                t.Rows[0].Cells[0].Paragraphs.First().AppendPicture(p1).Alignment = Alignment.center;
                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 生成用于放置宗地图的表格
        /// </summary>
        /// <param name="document"></param>
        /// <param name="countryname"></param>
        /// <param name="p1"></param>
        /// <returns></returns>
        public static Table parcelpicTabel(DocX document, Picture p1, string picname)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 1);
                t.Rows[0].Cells[0].Paragraphs.First().Append(picname).FontSize(22).Alignment = Alignment.center;
                t.Rows[1].Cells[0].Paragraphs.First().AppendPicture(p1).Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 添加用于放置鸟瞰图的表格
        /// </summary>
        /// <param name="document"></param>
        /// <param name="p_Orth"></param>
        /// <param name="p_front"></param>
        /// <param name="p_rear"></param>
        /// <param name="p_left"></param>
        /// <param name="p_right"></param>
        /// <returns></returns>
        public static Table AerialviewpicTable(DocX document, Picture p_Orth, Picture p_front, Picture p_rear, Picture p_left, Picture p_right)
        {
            Table t = null;
            try
            {
                t = document.AddTable(3, 2);

                t.Rows[0].Cells[0].Paragraphs.First().AppendPicture(p_Orth).Alignment = Alignment.center;//正射
                t.Rows[0].Cells[1].Paragraphs.First().AppendPicture(p_front).Alignment = Alignment.center;//前
                t.Rows[1].Cells[0].Paragraphs.First().AppendPicture(p_rear).Alignment = Alignment.center;//后
                t.Rows[1].Cells[1].Paragraphs.First().AppendPicture(p_left).Alignment = Alignment.center;//左
                t.Rows[2].Cells[0].Paragraphs.First().AppendPicture(p_right).Alignment = Alignment.center;//右

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[0].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[1].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[2].Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 生成用于放置房产外墙面实景图的表格
        /// </summary>
        /// <param name="document"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static Table picTable(DocX document, Picture p, string tblname)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 1);
                t.Rows[0].Cells[0].Paragraphs.First().Append(tblname).FontSize(16).Alignment = Alignment.center;
                t.Rows[1].Cells[0].Paragraphs.First().AppendPicture(p).Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Center;


            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }

        /// <summary>
        /// 生成用于放置管网平面图的表格
        /// </summary>
        /// <param name="document"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public static Table pipelinepicTable(DocX document, Picture p)
        {
            Table t = null;
            try
            {
                t = document.AddTable(2, 1);
                t.Rows[0].Cells[0].Paragraphs.First().Append("管网平面图").FontSize(16).Alignment = Alignment.center;
                t.Rows[1].Cells[0].Paragraphs.First().AppendPicture(p).Alignment = Alignment.center;

                t.Rows[0].Cells[0].VerticalAlignment = VerticalAlignment.Center;
                t.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Center;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(tableHelper), ex);
            }
            return t;
        }
    }
}
