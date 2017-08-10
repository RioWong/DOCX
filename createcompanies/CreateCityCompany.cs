using CreateWord.DB;
using CreateWord.listener;
using CreateWord.log;
using CreateWord.model;
using CreateWord.table;
using CreateWord.title;
using Novacode;
using RealEstate.Logic;
using RealEstate.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.createcompanies
{
    //生成一个市公司文档
    public class CreateCityCompany : CreateCompany
    {
        //string WordPath; //文档路径
        List<FDXXtbl_city> lstFM;
        NumoftitleHelper title_city = new NumoftitleHelper();//生成市公司文档时所使用的标题序号

        public CreateCityCompany() { }
        public CreateCityCompany(IDocCompilationListener docCompilationListener)
            : base(docCompilationListener)
        {            

        }

        /// <summary>
        /// 创建市公司文档
        /// </summary>
        public override void createword(string wordpath)
        {
            try
            {
                lstFM = FDXXtbl_city.GetInfo(CompanyID);
                //WordPath = wordpath;
                using (DocX document = DocX.Create(wordpath))
                {
                    setdoc(document);//设置文档属性

                    Addcover(document);//添加封面

                    Addtoc(document);//添加目录

                    Addintro(document,title_city);//添加概述

                    Addorganization(document, title_city);//添加组织机构

                    Addfdxx(document);//添加市公司的房地信息总汇

                    Addcountrycompanies(document);//添加市公司本部和县级公司

                    document.Save();

                    if (docCompilationListener != null)
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
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
                if (docCompilationListener != null)
                {
                    docCompilationListener.DocCompleted(new DocCompilationArg(
                        CompanyID, wordpath, DocCompilationStatus.Fail, ex.Message));
                }
                throw new Exception("生成失败：" + ex.Message);
            }

        }

        /// <summary>
        /// 创建市公司文档（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcityName"></param>
        /// <param name="childcityID"></param>
        public void createword(DocX document, string childcityName, int childcityID,NumoftitleHelper title)
        {
            try
            {
                lstFM = FDXXtbl_city.GetInfo(childcityID);

                Addcover(document, childcityName);//添加封面

                Addintro(document, childcityName, childcityID, title);//添加概述

                Addorganization(document, childcityID, title);//添加组织机构

                Addfdxx(document, childcityName, title);//添加房地信息总汇

                Addcountrycompanies(document, childcityID, title);//添加市公司本部和县级公司
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }
        }

        /// <summary>
        /// 添加封面（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcityName"></param>
        public void Addcover(DocX document, string childcityName)
        {
            int i = 0;
            try
            {
                Paragraph blank1 = document.InsertParagraph();
                blank1.InsertPageBreakBeforeSelf();
                while (i < 13) { i++; blank1.AppendLine(); }//空13行

                using( FontFamily fontFamily = new FontFamily("微软雅黑"))
                {
                    //标题
                    Paragraph title = document.InsertParagraph(childcityName + "非生产性房产资源汇编", false,
                        format.formatHelper.SetParagraphFormat(fontFamily, 48, Color.Black, true));
                    title.Alignment = Alignment.center;

                    Paragraph blank2 = document.InsertParagraph();
                    i = 0;
                    while (i < 30) { i++; blank2.AppendLine(); }//空13行
                    //日期
                    Paragraph _date = document.InsertParagraph(NumberToChinese(DateTime.Now.Year) + "年" + NumberToChinese(DateTime.Now.Month) + "月", false, 
                        format.formatHelper.SetParagraphFormat(fontFamily, 22, Color.Black, true));
                    _date.Alignment = Alignment.center;
                    document.DifferentFirstPage = true;
                    _date.InsertPageBreakAfterSelf();
                }
                

            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }
        }

        ///// <summary>
        ///// 添加概述模块
        ///// </summary>
        ///// <param name="document"></param>
        //public override void Addintro(DocX document)
        //{
        //    try
        //    {
        //        var h1 = document.InsertParagraph(CompanyName);
        //        h1.StyleName = "Heading1";
        //        var h1_1 = document.InsertParagraph("概述");
        //        h1_1.StyleName = "Heading2";

        //        string s = txt.txtHelper.readtxt(path.pathHelper.GetIntrotxtPath(CompanyID));
        //        Paragraph p = document.InsertParagraph(s);
        //        p.Font(new FontFamily("宋体")).FontSize(14);

        //        Picture p1 = picture.picHelper.getPic(document, path.pathHelper.GetIntropicPath(CompanyID), 330, 650);
        //        Paragraph pic = document.InsertParagraph();
        //        pic.AppendPicture(p1).Alignment = Alignment.center;
        //    }
        //    catch (System.Exception ex)
        //    {
        //        LogHelper.WriteLog(typeof(CreateCityCompany), ex);
        //    }
        //}

        /// <summary>
        /// 添加概述（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcityName"></param>
        /// <param name="childcountryID"></param>
        public void Addintro(DocX document, string childcityName, int childcountryID,NumoftitleHelper title)
        {
            try
            {
                string s = txt.txtHelper.readtxt(PathManager.getSingleton().GetIntrotxtPath(childcountryID, false));
                if (s == "") return;    //没有概述文件，则跳过
                title.Less1Zero();
                var h1 = document.InsertParagraph(title.num1title() + childcityName);
                h1.StyleName = "Heading1";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                }
               
                title.Less2Zero();
                var h1_1 = document.InsertParagraph(title.num2title() + "概述");
                h1_1.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_1.Color(Color.Black).FontSize(16).Font(fontfamily);
                }               
                Paragraph p = document.InsertParagraph(s);
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    p.Font(fontfamily).FontSize(14);
                }


                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetIntropicPath(childcountryID, false), 330, 650);
                Paragraph pic = document.InsertParagraph();
                pic.AppendPicture(p1).Alignment = Alignment.center;
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }
        }

        ///// <summary>
        ///// 添加市公司的组织机构模块
        ///// </summary>
        ///// <param name="document"></param>
        //public override void Addorganization(DocX document)
        //{
        //    var h1_2 = document.InsertParagraph("组织机构");
        //    h1_2.StyleName = "Heading2";
        //    Table t = table.tableHelper.organizationTable(document, Childcompany, path.pathHelper.GetOrganizationpicPath(CompanyID));
        //    document.InsertTable(t);
        //}

        /// <summary>
        /// 添加组织机构模块（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcityID"></param>
        public void Addorganization(DocX document, int childcityID, NumoftitleHelper title)
        {

            List<Companymodel> cchildcityID = DBhelper.GetChildcompany(childcityID);

            Table t = table.tableHelper.organizationTable(document, cchildcityID, PathManager.getSingleton().GetOrganizationpicPath(childcityID, false));//childcountryID
            t.Alignment = Alignment.center;
            t.AutoFit = AutoFit.Contents;
            if (t == null) return;

            title.Less2Zero();
            var h1_2 = document.InsertParagraph(title.num2title() + "组织机构");
            h1_2.InsertPageBreakBeforeSelf();
            h1_2.StyleName = "Heading2";
            using (FontFamily fontfamily = new FontFamily("宋体"))
            {
                h1_2.Color(Color.Black).FontSize(16).Font(fontfamily);
            }
            
            document.InsertTable(t);
        }

        /// <summary>
        /// 添加市公司的房地信息总汇模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addfdxx(DocX document)
        {
            try
            {
                title_city.Less2Zero();
                var h1_3 = document.InsertParagraph(title_city.num2title() + "房地信息总汇");
                h1_3.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_3.Color(Color.Black).FontSize(16).Font(fontfamily);
                }
                //表格描述
                var tbltitle = document.InsertParagraph(CompanyName + "房地信息汇总表");
                tbltitle.FontSize(14).Alignment = Alignment.center;
                Table t = tableHelper.Template_city(document);
                t = tableHelper.inserttable_city(t, lstFM);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                document.InsertTable(t);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }

        }

        /// <summary>
        /// 添加房地信息总汇模块（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcityName"></param>
        public void Addfdxx(DocX document, string childcityName, NumoftitleHelper title)
        {
            try
            {
                Table t = tableHelper.Template_city(document);
                t = tableHelper.inserttable_city(t, lstFM);
                if (t == null) return;
                title.Less2Zero();
                var h1_3 = document.InsertParagraph(title.num2title() + "房地信息总汇");
                h1_3.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_3.Color(Color.Black).FontSize(16).Font(fontfamily);
                }
                //表格描述
                var tbltitle = document.InsertParagraph(childcityName + "房地信息汇总表");
                tbltitle.FontSize(14).Alignment = Alignment.center;
         
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                document.InsertTable(t);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }

        }

        /// <summary>
        /// 添加下属县级公司模块
        /// </summary>
        /// <param name="document"></param>
        public override void Addcountrycompanies(DocX document)
        {
            try
            {
                foreach (Companymodel cm in Childcompany)//让市公司本部优先生成
                {
                    if (cm.property == "本部")
                    {
                        
                        title_city.Less1Zero();
                        var h1 = document.InsertParagraph(title_city.num1title() + cm.name);
                        h1.StyleName = "Heading1";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                        }
                        h1.InsertPageBreakBeforeSelf();
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, cm, title_city);
                        break;
                    }
                }
                foreach (Companymodel cm in Childcompany)
                {
                    if (cm.property != "本部")
                    {
                        title_city.Less1Zero();
                        var h1 = document.InsertParagraph(title_city.num1title() + cm.name);
                        h1.StyleName = "Heading1";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                        }
                        h1.InsertPageBreakBeforeSelf();
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, cm, title_city);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }
        }

        /// <summary>
        /// 添加市公司本部和县级公司模块（作为省公司下级的市公司）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="childcountryID"></param>
        public void Addcountrycompanies(DocX document, int childcountryID, NumoftitleHelper title)
        {
            try
            {
                List<Companymodel> CChildcompany = DBhelper.GetChildcompany(childcountryID);
                foreach (Companymodel ccm in CChildcompany)//让市公司本部优先生成
                {
                    if (ccm.property == "本部")
                    {
                        title.Less1Zero();
                        var h1 = document.InsertParagraph(title.num1title() + ccm.name);
                        h1.StyleName = "Heading1";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                        }
                        h1.InsertPageBreakBeforeSelf();
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, ccm, title);
                        break;
                    }
                }
                foreach (Companymodel ccm in CChildcompany)
                {
                    if (ccm.property != "本部")
                    {
                        title.Less1Zero();
                        var h1 = document.InsertParagraph(title.num1title() + ccm.name);
                        h1.StyleName = "Heading1";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                        }
                        h1.InsertPageBreakBeforeSelf();
                        CreateCountryCompany ccc = new CreateCountryCompany();
                        ccc.createword(document, ccm, title);
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCityCompany), ex);
            }
        }
    }

    
}
