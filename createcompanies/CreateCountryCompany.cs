using CreateWord.listener;
using CreateWord.log;
using CreateWord.model;
using CreateWord.parcels;
using CreateWord.table;
using CreateWord.title;
using Novacode;
using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.createcompanies
{
    /// <summary>
    /// 创建县公司文档类
    /// </summary>
    public class CreateCountryCompany : CreateCompany
    {
        //string WordPath; //文档输出路径
        List<FDXXtbl_country> lstFM;
        FDXXtbl_country FM;
        List<Parcelmodel> lstPM;
        List<Buildingmodel> lstBM;
        NumoftitleHelper title_country = new NumoftitleHelper();
        
        //string CompanyName = "";
        //int CompanyID = 0;

        //public CreateCountryCompany(int CountryCompanyID,string CountryCompanyName)
        //{
        //    CompanyName = CountryCompanyName;
        //    CompanyID = CountryCompanyID;
        //}

        public CreateCountryCompany() { }
        public CreateCountryCompany(IDocCompilationListener docCompilationListener)
            : base(docCompilationListener)
        {

        }

        /// <summary>
        /// 创建县公司文档
        /// </summary>
        public override void createword(string wordpath)
        {
            try
            {
                lstFM = FDXXtbl_country.GetInfo(CompanyID);
                FM = FDXXtbl_country.GetTotalInfo(CompanyID);
                lstPM = FDXXtbl_country.Parcels(lstFM);
                lstBM = FDXXtbl_country.Buildings(lstFM);
                //WordPath = wordpath;
                using (DocX document = DocX.Create(wordpath))
                {
                    //document.ApplyTemplate(@"F:\study\国网江苏\svnRealEstate\RealEstate.Web\国网江苏省电力公司\word.dotx");//应用模板

                    setdoc(document);//设置文档属性

                    Addcover(document);//添加封面

                    Addtoc(document);//添加目录

                    Addintro(document, title_country);//添加概述

                    Addloc(document);//添加位置分布图

                    Addfdxx(document);//添加房地信息统计

                    Addparcels(document);//添加宗地描述

                    document.Save();

                    if (docCompilationListener != null)
                    {
                        docCompilationListener.DocCompleted(new DocCompilationArg(
                            CompanyID, wordpath, DocCompilationStatus.Success));
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
                if (docCompilationListener != null)
                {
                    docCompilationListener.DocCompleted(new DocCompilationArg(
                        CompanyID, wordpath, DocCompilationStatus.Fail, ex.Message));
                }
                throw new Exception("生成失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 创建县公司文档（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryName"></param>
        /// <param name="childcountryID"></param>
        public void createword(DocX document, Companymodel childcountry,NumoftitleHelper title)
        {
            lstFM = FDXXtbl_country.GetInfo((int)childcountry.ID);
            FM = FDXXtbl_country.GetTotalInfo((int)childcountry.ID);
            lstPM = FDXXtbl_country.Parcels(lstFM);
            lstBM = FDXXtbl_country.Buildings(lstFM);

            if (!childcountry.name.Contains("本部"))
                Addintro(document, childcountry, title);   //不是市公司本部，则添加概述
            Addloc(document, childcountry, title);//添加位置分布图
            Addfdxx(document, childcountry, title);//添加房地信息统计
            Addparcels(document, childcountry, title);//添加宗地描述

            document.Save();
        }


        #region 各个模块


        /// <summary>
        /// 添加概述模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addintro(DocX document, NumoftitleHelper title)
        {
            try
            {
                title.Less1Zero();
                var h1 = document.InsertParagraph(title.num1title() + "概述");
                h1.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                }

                string s = txt.txtHelper.readtxt(PathManager.getSingleton().GetIntrotxtPath(CompanyID, false));
                Paragraph p = document.InsertParagraph(s);               
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    p.Font(fontfamily).FontSize(14);
                }
                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetIntropicPath(CompanyID, false), 330, 650);
                Paragraph pic = document.InsertParagraph();
                pic.AppendPicture(p1).Alignment = Alignment.center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加概述模块（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryID"></param>
        /// <param name="ischild"></param>
        public void Addintro(DocX document, Companymodel childcountry,NumoftitleHelper title)
        {
            try
            {
                Paragraph h1;
                string s = txt.txtHelper.readtxt(PathManager.getSingleton().GetIntrotxtPath((int)childcountry.ID, false));
                if (s == "") return;    //概述文件缺少，就不添加概述
                if (childcountry.property == "直属单位"||childcountry.property =="培训单位")
                {
                    title.Less3Zero();
                    h1 = document.InsertParagraph(title.num3title() + "概述");
                    h1.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else
                {
                    title.Less2Zero();
                    h1 = document.InsertParagraph(title.num2title() + "概述");
                    h1.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                }
              
                Paragraph p = document.InsertParagraph(s);
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    p.Font(fontfamily).FontSize(14);
                }
                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetIntropicPath((int)childcountry.ID, false), 330, 650);
                Paragraph pic = document.InsertParagraph();
                pic.AppendPicture(p1).Alignment = Alignment.center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加位置分布图模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addloc(DocX document)
        {
            try
            {
                title_country.Less1Zero();
                var h2 = document.InsertParagraph(title_country.num1title() + "位置分布图");
                // h1.Font(new FontFamily("宋体")).FontSize(16).Bold();
                h2.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h2.Color(Color.Black).FontSize(22).Font(fontfamily);
                }
                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetLocpicPath(CompanyID, false), 450, 886);
                //Paragraph pic = document.InsertParagraph();
                //pic.AppendPicture(p1).Alignment = Alignment.center;

                var tbl_deed = document.InsertParagraph();//房产证表格
                Table t_deed = tableHelper.picTable(document, p1, "位置分布图");
                t_deed.Alignment = Alignment.center;
                t_deed.AutoFit = AutoFit.Contents;
                tbl_deed.InsertTableAfterSelf(t_deed);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加位置分布图模块（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryID"></param>
        public void Addloc(DocX document, Companymodel childcountry,NumoftitleHelper title)
        {
            try
            {
                Paragraph h2;
                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetLocpicPath((int)childcountry.ID, false), 384, 864);
                if (p1 == null) return; //如果图片文件不存在，则跳过这个模块
                if (childcountry.property == "直属单位" || childcountry.property == "培训单位")
                {
                    
                    title.Less3Zero();
                    h2 = document.InsertParagraph(title.num3title() + "位置分布图");
                    h2.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h2.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }                   
                }
                else
                {
                    title_country.Less2Zero();
                    h2 = document.InsertParagraph(title.num2title() + "位置分布图");
                    h2.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h2.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                    
                }
                Table t = tableHelper.locpicTable(document, p1);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                Paragraph pic = document.InsertParagraph();
                pic.InsertTableAfterSelf(t).Alignment = Alignment.center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加房地信息统计模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addfdxx(DocX document)
        {
            try
            {
                title_country.Less1Zero();
                var h3 = document.InsertParagraph(title_country.num1title() + "房地信息统计");
                //h3.Font(new FontFamily("宋体")).FontSize(16).Bold();
                h3.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h3.Color(Color.Black).FontSize(22).Font(fontfamily);
                }
                //文字描述
                var p = document.InsertParagraph();
                p.Append("    " + CompanyName + "市公司现有各类用房");
                p.AppendBookmark("各类用房栋数");
                p.Append("栋，占地总面积");
                p.AppendBookmark("占地总面积");
                p.Append("平方米,总建筑面积");
                p.AppendBookmark("总建筑面积");
                p.Append("平方米。其中");
                p.AppendBookmark("各类用房面积");
                p.Append("；建成投运10年内的房屋面积为");
                p.AppendBookmark("十年内房屋面积");
                p.Append("平方米，建成投运10-20年的房屋面积为");
                p.AppendBookmark("十到二十年内房屋面积");
                p.Append("平方米，建成投运20-30年的房屋面积为");
                p.AppendBookmark("二十到三十年内房屋面积");
                p.Append("平方米，建成投运30年以上的房屋面积为");
                p.AppendBookmark("三十年以上房屋面积");
                p.Append("平方米。");
                finishBM(document);//完成书签内容
                //表格描述
                var tbltitle = document.InsertParagraph("房地信息汇总表");
                tbltitle.FontSize(14).Alignment = Alignment.center;
                Table t = tableHelper.Template_country(document);
                t = tableHelper.inserttable_country(t, lstFM, FM);
                t = tableHelper.combineCells(t, lstPM, lstBM);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                document.InsertTable(t);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加房地信息统计模块（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryName"></param>
        /// <param name="childcountryID"></param>
        public void Addfdxx(DocX document, Companymodel childcountry,NumoftitleHelper title)
        {
            try
            {
                Paragraph h3;

                if (childcountry.property == "直属单位" || childcountry.property == "培训单位") 
                {
                    title.Less3Zero();
                    h3 = document.InsertParagraph(title.num3title() + "房地信息统计");
                    h3.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h3.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                  
                }
                else
                {
                    title.Less2Zero();
                    h3 = document.InsertParagraph(title.num2title() + "房地信息统计");
                    h3.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h3.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                }
                //文字描述
                var p = document.InsertParagraph();
                p.Append(childcountry.name + "市公司现有各类用房");
                p.AppendBookmark((int)childcountry.ID + "各类用房栋数");
                p.Append("栋，占地总面积");
                p.AppendBookmark((int)childcountry.ID + "占地总面积");
                p.Append("平方米,总建筑面积");
                p.AppendBookmark((int)childcountry.ID + "总建筑面积");
                p.Append("平方米。其中");
                p.AppendBookmark((int)childcountry.ID + "各类用房面积");
                p.Append("；建成投运10年内的房屋面积为");
                p.AppendBookmark((int)childcountry.ID + "十年内房屋面积");
                p.Append("平方米，建成投运10-20年的房屋面积为");
                p.AppendBookmark((int)childcountry.ID + "十到二十年内房屋面积");
                p.Append("平方米，建成投运20-30年的房屋面积为");
                p.AppendBookmark((int)childcountry.ID + "二十到三十年内房屋面积");
                p.Append("平方米，建成投运30年以上的房屋面积为");
                p.AppendBookmark((int)childcountry.ID + "三十年以上房屋面积");
                p.Append("平方米。");
                finishBM(document, (int)childcountry.ID);//完成书签内容

                //表格描述
                var tbltitle = document.InsertParagraph("房地信息汇总表");
                tbltitle.FontSize(14).Alignment = Alignment.center;
                Table t = tableHelper.Template_country(document);
                t = tableHelper.inserttable_country(t, lstFM, FM);
                t = tableHelper.combineCells(t, lstPM, lstBM);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                document.InsertTable(t);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }

        /// <summary>
        /// 添加宗地描述模块
        /// </summary>
        /// <param name="document"></param>
        public void Addparcels(DocX document)
        {
            foreach (Parcelmodel pm in lstPM)
            {
                parcelHelper phelper = new parcelHelper(pm, CompanyID);//添加各个地块的信息
                phelper.insertInfo(document, lstFM, lstPM, false, false, title_country);
            }

        }

        /// <summary>
        /// 添加宗地描述模块（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryID"></param>
        public void Addparcels(DocX document, Companymodel childcountry, NumoftitleHelper title)
        {
            foreach (Parcelmodel pm in lstPM)
            {
                parcelHelper phelper = new parcelHelper(pm, (int)childcountry.ID);//添加各个地块的信息
                if (childcountry.property == "培训单位" || childcountry.property == "直属单位")
                    phelper.insertInfo(document, lstFM, lstPM, false, true, title);
                else
                    phelper.insertInfo(document, lstFM, lstPM, true, false, title);
            }
        }
        #endregion



        /// <summary>
        /// 完成书签内容
        /// </summary>
        /// <param name="document"></param>
        public void finishBM(DocX document)
        {
            countryFDXXsentence1 fs1 = new countryFDXXsentence1();
            List<countryFDXXsentence2> lstFS2 = new List<countryFDXXsentence2>();
            countryFDXXsentence3 fs3 = new countryFDXXsentence3();
            string temp = "";
            try
            {
                fs1 = countryFDXXsentence1.GetInfo(CompanyID);
                lstFS2 = countryFDXXsentence2.GetInfo(CompanyID);
                fs3 = countryFDXXsentence3.GetInfo(CompanyID);
                document.Bookmarks["各类用房栋数"].SetText("" + fs1.count);
                //document.Bookmarks["各类用房栋数"].Paragraph.Append("" + fs1.count).FontSize(14);
                document.Bookmarks["占地总面积"].SetText("" + fs1.ZDZMJ);
                //document.Bookmarks["占地总面积"].Paragraph.Append("" + fs1.ZDZMJ).FontSize(14);
                document.Bookmarks["总建筑面积"].SetText("" + fs1.ZJZMJ);
                //document.Bookmarks["总建筑面积"].Paragraph.Append("" + fs1.ZJZMJ).FontSize(14);

                foreach (countryFDXXsentence2 fs2 in lstFS2)
                {
                    temp += fs2.GNGL;
                    temp += "" + fs2.GNGL_MJ;
                    temp += "平方米，";
                }
                temp = temp.Trim('，');
                document.Bookmarks["各类用房面积"].SetText(temp);
                //document.Bookmarks["各类用房面积"].Paragraph.Append(temp).FontSize(14);
                document.Bookmarks["十年内房屋面积"].SetText("" + fs3.FWMJ_10);
                //document.Bookmarks["十年内房屋面积"].Paragraph.Append("" + fs3.FWMJ_10).FontSize(14);
                document.Bookmarks["十到二十年内房屋面积"].SetText("" + fs3.FWMJ_1020);
                //document.Bookmarks["十到二十年内房屋面积"].Paragraph.Append("" + fs3.FWMJ_20).FontSize(14);
                document.Bookmarks["二十到三十年内房屋面积"].SetText("" + fs3.FWMJ_2030);
                //document.Bookmarks["十到二十年内房屋面积"].Paragraph.Append("" + fs3.FWMJ_20).FontSize(14);
                document.Bookmarks["三十年以上房屋面积"].SetText("" + fs3.FWMJ_30);
                //document.Bookmarks["三十年以上房屋面积"].Paragraph.Append("" + fs3.FWMJ_30).FontSize(14);
                document.Bookmarks["三十年以上房屋面积"].Paragraph.FontSize(14);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }
        /// <summary>
        /// 完成书签内容（作为市公司的下级）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryID"></param>
        public void finishBM(DocX document, int childcountryID)
        {
            countryFDXXsentence1 fs1 = new countryFDXXsentence1();
            List<countryFDXXsentence2> lstFS2 = new List<countryFDXXsentence2>();
            countryFDXXsentence3 fs3 = new countryFDXXsentence3();
            string temp = "";
            try
            {
                fs1 = countryFDXXsentence1.GetInfo(childcountryID);
                lstFS2 = countryFDXXsentence2.GetInfo(childcountryID);
                fs3 = countryFDXXsentence3.GetInfo(childcountryID);
                document.Bookmarks[childcountryID + "各类用房栋数"].SetText("" + fs1.count);
                document.Bookmarks[childcountryID + "占地总面积"].SetText("" + fs1.ZDZMJ);
                document.Bookmarks[childcountryID + "总建筑面积"].SetText("" + fs1.ZJZMJ);

                foreach (countryFDXXsentence2 fs2 in lstFS2)
                {
                    temp += fs2.GNGL;
                    temp += "" + fs2.GNGL_MJ;
                    temp += "平方米，";
                }
                temp = temp.Trim('，');
                document.Bookmarks[childcountryID + "各类用房面积"].SetText(temp);
                document.Bookmarks[childcountryID + "十年内房屋面积"].SetText("" + fs3.FWMJ_10);
                document.Bookmarks[childcountryID + "十到二十年内房屋面积"].SetText("" + fs3.FWMJ_1020);
                document.Bookmarks[childcountryID + "二十到三十年内房屋面积"].SetText("" + fs3.FWMJ_2030);
                document.Bookmarks[childcountryID + "三十年以上房屋面积"].SetText("" + fs3.FWMJ_30);
                document.Bookmarks[childcountryID + "三十年以上房屋面积"].Paragraph.FontSize(14);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
        }
    }
}
