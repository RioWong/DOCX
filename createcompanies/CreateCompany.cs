using CreateWord.listener;
using CreateWord.log;
using CreateWord.model;
using CreateWord.title;
using Novacode;
using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace CreateWord.createcompanies
{
    /// <summary>
    /// 创建公司文档类
    /// </summary>
    public class CreateCompany
    {
        private string _CompanyName = "";//公司名称
        private int _CompanyID = 0;//公司ID
        private List<Companymodel> childcompany = new List<Companymodel>();
        protected IDocCompilationListener docCompilationListener = null;
   

        public string CompanyName
        {
            get { return _CompanyName; }
            set { _CompanyName = value; }
        }
        public int CompanyID
        {
            get { return _CompanyID; }
            set { _CompanyID = value; }
        }
        public List<Companymodel> Childcompany
        {
            get { return childcompany; }
            set { childcompany = value; }
        }

        public CreateCompany(){ }

        public CreateCompany( IDocCompilationListener docCompilationListener )
        {
            this.docCompilationListener = docCompilationListener;
        }

        /// <summary>
        /// 创建相应公司的文档
        /// </summary>
        public virtual void createword(string wordpath) { }


        /// <summary>
        /// 设置文档属性
        /// </summary>
        /// <param name="document"></param>
        public virtual void setdoc(DocX document)
        {
            //document.AddHeaders();
            document.PageWidth = 831.6f;
            document.PageHeight = 1176f;
        }

        /// <summary>
        /// 添加封面
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addcover(DocX document)
        {
            int i = 0;
            
            try
            {
                Paragraph blank1 = document.InsertParagraph();
                while (i < 13) { i++; blank1.AppendLine(); }//空13行

                using( FontFamily fontFamily = new FontFamily("微软雅黑"))
                {
                    //标题
                    Paragraph title = document.InsertParagraph(CompanyName + "非生产性房产资源汇编", false, format.formatHelper.SetParagraphFormat(fontFamily, 48, Color.Black, true));
                    title.Alignment = Alignment.center;
                }
                

                //日期
                Paragraph _date = document.InsertParagraph();
                i = 0;
                while (i < 30) { i++; _date.AppendLine(); }//空13行
                _date.Append(NumberToChinese(DateTime.Now.Year) + "年" + NumberToChinese(DateTime.Now.Month) + "月")
                    .FontSize(22)
                    .Bold()
                    .Alignment = Alignment.center;
                document.DifferentFirstPage = true;
                _date.InsertPageBreakAfterSelf();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCompany), ex);
            }

        }

        /// <summary>
        /// 添加目录
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addtoc(DocX document)
        {
            try
            {
                Formatting f = new Formatting();
                f.Size = 22;
                Paragraph p = document.InsertParagraph("目     录", false, f).Bold();
                p.Alignment = Alignment.center;
                Paragraph toc = document.InsertParagraph();
                document.InsertTableOfContents(toc, "", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2", 4);
                toc.InsertPageBreakAfterSelf();
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCompany), ex);
            }

        }

        /// <summary>
        /// 添加概述模块（用于市公司）
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addintro(DocX document,NumoftitleHelper title)
        {
            try
            {
                title.Less1Zero();
                Paragraph h1 = document.InsertParagraph(title.num1title() + CompanyName);
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

                string s = txt.txtHelper.readtxt(PathManager.getSingleton().GetIntrotxtPath(CompanyID, false));
                Paragraph p = document.InsertParagraph(s).FontSize(14);

                Picture p1 = picture.picHelper.getPic(document, PathManager.getSingleton().GetIntropicPath(CompanyID, false), 330, 650);
                if( p1 != null )
                {
                    Paragraph pic = document.InsertParagraph();
                    pic.AppendPicture(p1).Alignment = Alignment.center;
                }                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCompany), ex);
            }
        }

        ///// <summary>
        ///// 添加概述模块（用于县公司）
        ///// </summary>
        ///// <param name="document"></param>
        ///// <param name="ischild"></param>
        //public virtual void Addintro(DocX document,bool ischild) { }


        /// <summary>
        /// 添加组织机构模块（省公司和市公司）
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addorganization(DocX document, NumoftitleHelper title)
        {
            try
            {
                Paragraph pagebreak = document.InsertParagraph();
                pagebreak.InsertPageBreakAfterSelf();//分页符
                title.Less2Zero();
                var h1_2 = document.InsertParagraph(title.num2title() + "组织机构");
                h1_2.StyleName = "Heading2";
                using (FontFamily fontfamily = new FontFamily("宋体"))
                {
                    h1_2.Color(Color.Black).FontSize(16).Font(fontfamily);
                }
                Table t = table.tableHelper.organizationTable(document, Childcompany, PathManager.getSingleton().GetOrganizationpicPath(CompanyID, false));
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                if (t != null) document.InsertTable(t);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCompany), ex);
            }
        }

        /// <summary>
        /// 添加位置分布图模块
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addloc(DocX document) { }

        /// <summary>
        /// 添加房地信息模块
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addfdxx(DocX document) { }

        /// <summary>
        /// 添加公司本部
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addheadquarters(DocX document) { }

        /// <summary>
        /// 添加县级公司模块
        /// </summary>
        /// <param name="document"></param>
        public virtual void Addcountrycompanies(DocX document) { }
     
        /// <summary>
        /// 阿拉伯数字变中文数字
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public string NumberToChinese(int number)
        {
            string outString = string.Empty;
            char d;
            try
            {


                if (number > 12)//年份
                {
                    string year = number.ToString();
                    foreach (char s in year)
                    {
                        d = ' ';
                        switch (s)
                        {
                            case '0': d = '零'; break;
                            case '1': d = '一'; break;
                            case '2': d = '二'; break;
                            case '3': d = '三'; break;
                            case '4': d = '四'; break;
                            case '5': d = '五'; break;
                            case '6': d = '六'; break;
                            case '7': d = '七'; break;
                            case '8': d = '八'; break;
                            case '9': d = '九'; break;
                        }
                        outString += d;
                    }
                }
                else//月份
                {
                    switch (number)
                    {
                        case 1: outString = "一"; break;
                        case 2: outString = "二"; break;
                        case 3: outString = "三"; break;
                        case 4: outString = "四"; break;
                        case 5: outString = "五"; break;
                        case 6: outString = "六"; break;
                        case 7: outString = "七"; break;
                        case 8: outString = "八"; break;
                        case 9: outString = "九"; break;
                        case 10: outString = "十"; break;
                        case 11: outString = "十一"; break;
                        case 12: outString = "十二"; break;
                    }
                }

            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(CreateCountryCompany), ex);
            }
            return outString;
        }
    }
}
