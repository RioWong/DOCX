using CreateWord.DB;
using CreateWord.listener;
using CreateWord.log;
using CreateWord.model;
using CreateWord.table;
using CreateWord.title;
using Novacode;
using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;

namespace CreateWord.createcompanies
{
    
    //生成省公司文档
    public class CreateProvinceCompany : CreateCompany
    {
        //string WordPath; //文档路径
        List<FDXXtbl_province_summary> lstFM_summary;
        List<FDXXtbl_province_analysis> lstFM_analysis;
        NumoftitleHelper title_province = new NumoftitleHelper();//生成省公司文档时所使用的标题序号

        public CreateProvinceCompany() { }
        public CreateProvinceCompany(IDocCompilationListener docCompilationListener)
            : base(docCompilationListener)
        {
        }
        
        /// <summary>
        /// 创建省公司文档
        /// </summary>
        public override void createword(string wordpath)
        {
            //WordPath = wordpath;
            lstFM_summary = FDXXtbl_province_summary.GetInfo();
            lstFM_analysis = FDXXtbl_province_analysis.GetInfo();

            AddProvincePart(wordpath);// 添加省公司单独的文档
            //AddCityPart();//添加省公司下面各个市级公司单独的文档
        }

        /// <summary>
        /// 添加省公司单独的文档
        /// </summary>
        public void AddProvincePart(string wordpath)
        {
            try
            {
                using (DocX document = DocX.Create(wordpath))
                {
                    setdoc(document);//设置文档属性

                    Addcover(document);//添加封面

                    Addtoc(document);//添加目录

                    Addintro(document, title_province);//添加概述

                    Addorganization(document, title_province);//添加组织机构

                    Addfdxx(document);//添加省公司的房地信息总汇

                    Addheadquarters(document);//添加省公司本部

                    Addsubordinates(document);//添加省公司直属单位

                    Addtrainings(document);//添加省公司培训单位

                    document.Save();

                    if( docCompilationListener != null )
                    {
                        docCompilationListener.DocCompleted(new DocCompilationArg(
                            CompanyID, wordpath, DocCompilationStatus.Success));
                    }
                }
                //Console.WriteLine("创建文档成功！");
                //Console.WriteLine(WordPath);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
                if (docCompilationListener != null)
                {
                    docCompilationListener.DocCompleted(new DocCompilationArg(
                        CompanyID, wordpath, DocCompilationStatus.Fail, ex.Message ));
                }
                throw new Exception("生成失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 添加省公司下面各个市级公司单独的文档
        /// </summary>
        public void AddCityPart()
        {
            try
            {
                IList<Companymodel> lstCM = DBhelper.GetChildcompany(CompanyID);
                var tasks = new Task[lstCM.Count];
                for (int i = 0; i < lstCM.Count; ++i )
                {
                    LogHelper.WriteLog(typeof(CreateProvinceCompany), "AddCityPart:" + lstCM[i].ID.ToString());
                    tasks[i] = Task.Run(() => (new CREATE( docCompilationListener)).compilationdocument((int)lstCM[i].ID));
                }
                Task.WaitAll(tasks);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }

        }

        /// <summary>
        /// 添加组织机构模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addorganization(DocX document, NumoftitleHelper title)
        {
            try
            {
                title_province.Less2Zero();
                var h1_2 = document.InsertParagraph(title.num2title() + CompanyName + "组织机构");
                h1_2.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_2.Color(Color.Black).FontSize(16).Font(fontfamily);
                }
                Paragraph pic = document.InsertParagraph();
                pic.InsertPicture(picture.picHelper.getPic(document, PathManager.getSingleton().GetOrganizationpicPath(CompanyID, false), 1306, 936)).Alignment = Alignment.center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }
        }

        /// <summary>
        /// 添加省公司的房地信息总汇模块
        /// </summary>
        public override void Addfdxx(DocX document)
        {

            try
            {
                title_province.Less2Zero();
                var h1_3 = document.InsertParagraph(title_province.num2title() + CompanyName + "江苏省房地信息总汇");
                h1_3.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_3.Color(Color.Black).FontSize(16).Font(fontfamily);
                }
                
                h1_3.InsertPageBreakBeforeSelf();
                //文字描述
                var p = document.InsertParagraph();
                p.Append("截至2015年6月30日，公司所属单位现有土地总面积约为");
                p.AppendBookmark("土地总面积");
                p.Append("亩，用房总面积约");
                p.AppendBookmark("用房总面积");
                p.Append("万平方米（不含在建项目，地上");
                p.AppendBookmark("地上面积");
                p.Append("万平方米，地下");
                p.AppendBookmark("地下面积");
                p.Append("万平方米）。其中：危房：");
                p.AppendBookmark("危房面积");
                p.Append("万平方米，占总面积");
                p.AppendBookmark("危房面积比");
                p.Append("；规划拆除用房：");
                p.AppendBookmark("规划拆除用房面积");
                p.Append("万平方米，占总面积");
                p.AppendBookmark("规划拆除用房面积比");
                p.Append("；未办权证用房：");
                p.AppendBookmark("未办权证用房面积");
                p.Append("万平方米，占总面积");
                p.AppendBookmark("未办权证用房面积比");
                p.Append("。其中2013年及以前房屋面积");
                p.AppendBookmark("二零一三年以前的房屋面积");
                p.Append("万平方米（占总用房面积");
                p.AppendBookmark("二零一三年以前的房屋面积比");
                p.Append("）。");
                finishBM(document);
                //表格描述
                var tbltitle_summary = document.InsertParagraph("江苏公司现有土地、房地总面积汇总表");
                tbltitle_summary.FontSize(14).Alignment = Alignment.center;
                Table t_summary = tableHelper.Template_province_summary(document);
                t_summary = tableHelper.inserttable_province_summary(t_summary, lstFM_summary);
                t_summary.Alignment = Alignment.center;
                t_summary.AutoFit = AutoFit.Contents;
                document.InsertTable(t_summary);

                var tbltitle_analysis = document.InsertParagraph("江苏公司企业用房分析表");
                tbltitle_analysis.SpacingBefore(15);
                tbltitle_analysis.FontSize(14).Alignment = Alignment.center;
                Table t_analysis = tableHelper.Template_province_analysis(document);
                t_analysis = tableHelper.inserttable_province_analysis(t_analysis, lstFM_analysis);
                t_analysis = tableHelper.combineCells(t_analysis);
                t_analysis.AutoFit = AutoFit.Contents;
                //t_analysis.AutoFit = AutoFit.Contents;
                t_analysis.Alignment = Alignment.center;
                document.InsertTable(t_analysis);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }
        }

        /// <summary>
        /// 添加省公司本部模块
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryName"></param>
        /// <param name="childcountryID"></param>
        public override void Addheadquarters(DocX document)
        {
            try
            {
                foreach (Companymodel cm in Childcompany)
                {
                    if (cm.property == "本部")
                    {
                        title_province.Less1Zero();
                        var h1 = document.InsertParagraph(title_province.num1title() + cm.name);
                        h1.InsertPageBreakBeforeSelf();//分页符
                        h1.StyleName = "Heading1";
                        using(FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                        }
                        CreateCountryCompany ccc = new CreateCountryCompany( );
                        ccc.createword(document, cm, title_province);
                        break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }

        }

        /// <summary>
        /// 添加省公司所属直属单位模块
        /// </summary>
        /// <param name="document"></param>
        public void Addsubordinates(DocX document)
        {
            try
            {
                title_province.Less1Zero();
                var h1 = document.InsertParagraph(title_province.num1title() + "省公司所属直属单位");
                h1.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                }                

                foreach (Companymodel cm in Childcompany)
                {
                    if (cm.property == "直属单位")
                    {
                        title_province.Less2Zero();
                        var h1_1 = document.InsertParagraph(title_province.num2title() + cm.name);
                        h1_1.StyleName = "Heading2";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1_1.Color(Color.Black).FontSize(16).Font(fontfamily);
                        }                        
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, cm, title_province);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }
        }

        /// <summary>
        /// 添加省公司培训单位模块
        /// </summary>
        /// <param name="document"></param>
        public void Addtrainings(DocX document)
        {
            try
            {
                title_province.Less1Zero();
                var h1 = document.InsertParagraph(title_province.num1title() + "省公司所属培训单位");
                h1.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                }
                

                foreach (Companymodel cm in Childcompany)
                {
                    if (cm.property == "培训单位")
                    {
                        title_province.Less2Zero();
                        var h1_1 = document.InsertParagraph(title_province.num2title() + cm.name);
                        h1_1.StyleName = "Heading2";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1_1.Color(Color.Black).FontSize(16).Font(fontfamily);
                        }                        
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, cm, title_province);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }
        }

        ///// <summary>
        ///// 添加市公司模块
        ///// </summary>
        ///// <param name="document"></param>
        //public void Addcitycompanies(DocX document)
        //{
        //    try
        //    {
        //        foreach (Companymodel cm in Childcompany)
        //        {
        //            if (cm.property == "供电公司")
        //            {
        //                title_province = new NumoftitleHelper();
        //                CreateCityCompany ccc = new CreateCityCompany();
        //                ccc.createword(document, cm.name, (int)cm.ID, title_province);
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
        //    }
        //}

        /// <summary>
        /// 完成书签内容
        /// </summary>
        /// <param name="document"></param>
        public void finishBM(DocX document)
        {
            provinceFDXXsentence1 fs1 = new provinceFDXXsentence1();
            provinceFDXXsentence2 fs2 = new provinceFDXXsentence2();
            try
            {
                fs1 = provinceFDXXsentence1.GetInfo();
                fs2 = provinceFDXXsentence2.GetInfo();
                //第一句中的书签
                document.Bookmarks["土地总面积"].SetText("" + fs1.ZDMJ);
                //document.Bookmarks["土地总面积"].Paragraph.Append("" + fs1.ZDMJ).FontSize(14);
                document.Bookmarks["用房总面积"].SetText("" + Math.Round(fs1.FCMJ / 10000, 2));
                //document.Bookmarks["用房总面积"].Paragraph.Append("" + Math.Round(fs1.FCMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["地上面积"].SetText("" + Math.Round(fs1.DSMJ / 10000, 2));
                //document.Bookmarks["地上面积"].Paragraph.Append("" + Math.Round(fs1.DSMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["地下面积"].SetText("" + Math.Round(fs1.DXMJ / 10000, 2));
                //document.Bookmarks["地下面积"].Paragraph.Append("" + Math.Round(fs1.DXMJ / 10000, 2)).FontSize(14);
                //第二句中的书签
                document.Bookmarks["危房面积"].SetText("" + Math.Round(fs2.WFZMJ / 10000, 2));
                //document.Bookmarks["危房面积"].Paragraph.Append("" + Math.Round(fs2.WFZMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["危房面积比"].SetText(fs2.WFZMJB);
                //document.Bookmarks["危房面积比"].Paragraph.Append(fs2.WFZMJB).FontSize(14);
                document.Bookmarks["规划拆除用房面积"].SetText("" + Math.Round(fs2.CCZMJ / 10000, 2));
                //document.Bookmarks["规划拆除用房面积"].Paragraph.Append("" + Math.Round(fs2.CCZMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["规划拆除用房面积比"].SetText(fs2.CCZMJB);
                //document.Bookmarks["规划拆除用房面积比"].Paragraph.Append(fs2.CCZMJB).FontSize(14);
                document.Bookmarks["未办权证用房面积"].SetText("" + Math.Round(fs2.NQZZMJ / 10000, 2));
                //document.Bookmarks["未办权证用房面积"].Paragraph.Append("" + Math.Round(fs2.NQZZMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["未办权证用房面积比"].SetText(fs2.NQZZMJB);
                //document.Bookmarks["未办权证用房面积比"].Paragraph.Append(fs2.NQZZMJB).FontSize(14);
                document.Bookmarks["二零一三年以前的房屋面积"].SetText("" + Math.Round(fs2.BF2013ZMJ / 10000, 2));
                //document.Bookmarks["二零一三年以前的房屋面积"].Paragraph.Append("" + Math.Round(fs2.BF2013ZMJ / 10000, 2)).FontSize(14);
                document.Bookmarks["二零一三年以前的房屋面积比"].SetText(fs2.BF2013ZMJB);
                //document.Bookmarks["二零一三年以前的房屋面积比"].Paragraph.Append(fs2.BF2013ZMJB).FontSize(14);
                document.Bookmarks["二零一三年以前的房屋面积比"].Paragraph.FontSize(14);     //可以设置整个段落的字体
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateProvinceCompany), ex);
            }
        }
    }
}
