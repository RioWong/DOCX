using CreateWord.log;
using CreateWord.model;
using CreateWord.picture;
using CreateWord.table;
using CreateWord.title;
using Novacode;
using RealEstate.Logic;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.parcels
{

    public class parcelHelper
    {
        private string parcelname = "";//宗地名称
        private int parcelID;//宗地ID
        private int companyID;//宗地所属公司的ID

        //public parcelHelper(string name, int id)
        public parcelHelper( Parcelmodel parcel, int companyID)
        {
            this.parcelname = parcel.name;
            this.parcelID = parcel.id;
            this.companyID = companyID ;
        }

        /// <summary>
        /// 插入某个宗地的所有信息（宗地图、平面分布图、鸟瞰图、分层分户平面图、场地涉外管线布置图）
        /// </summary>
        /// <param name="document"></param>
        /// <param name="lstFM"></param>
        /// <param name="p"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        /// <param name="istrainings">判断是否为培训单位</param>
        /// <param name="title"></param>
        public void insertInfo(DocX document, List<FDXXtbl_country> lstFM, List<Parcelmodel> p, bool ischild, bool istrainings, NumoftitleHelper title)
        {
            try
            {
                Paragraph pagebreak = document.InsertParagraph();
                pagebreak.InsertPageBreakAfterSelf();//分页符
                Paragraph h1;
                if (ischild)
                {
                    title.Less2Zero();
                    h1 = document.InsertParagraph(title.num2title() + parcelname);
                    h1.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }

                }
                else if (istrainings)
                {
                    title.Less3Zero();
                    h1 = document.InsertParagraph(title.num3title() + parcelname);
                    h1.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else
                {
                    title.Less1Zero();
                    h1 = document.InsertParagraph(title.num1title() + parcelname);
                    h1.StyleName = "Heading1";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1.Color(Color.Black).FontSize(22).Font(fontfamily);
                    }
                }
                insertZDT(document, ischild, istrainings, title);//宗地图
                insertZPMFBT(document, ischild, istrainings, title);//总平面分布图
                insertNKT(document, ischild, istrainings, title);//鸟瞰图
                insertFCFH(document, lstFM, p, ischild, istrainings, title);//分层分户平面图
                insertGXT(document, ischild, istrainings, title);//场地涉外管线布置图
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入宗地图模块
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        public void insertZDT(DocX document, bool ischild, bool istrainings, NumoftitleHelper title)
        {
            try
            {
                Paragraph h1_1;
                Picture p = picHelper.getPic(document, PathManager.getSingleton().GetParcelpicPath( parcelID , false), 724, 833);
                //if (p == null) return;
                if (p != null)
                {
                    if (ischild)
                    {
                        title.Less3Zero();
                        h1_1 = document.InsertParagraph(title.num3title() + "宗地图");
                        h1_1.StyleName = "Heading3";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1_1.Color(Color.Black).FontSize(14).Font(fontfamily);
                        }

                    }
                    else if (istrainings)
                    {
                        title.Less4Zero();
                        h1_1 = document.InsertParagraph(title.num4title() + "宗地图");
                        h1_1.StyleName = "Heading4";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1_1.Color(Color.Black).FontSize(14).Font(fontfamily);
                        }
                    }
                    else
                    {
                        title.Less2Zero();
                        h1_1 = document.InsertParagraph(title.num2title() + "宗地图");
                        h1_1.StyleName = "Heading2";
                        using (FontFamily fontfamily = new FontFamily("宋体"))
                        {
                            h1_1.Color(Color.Black).FontSize(16).Font(fontfamily);
                        }
                    }
                    h1_1.AppendLine();
                    var parcelPic = document.InsertParagraph();
                    Table t = tableHelper.parcelpicTabel(document, p, "宗地图");
                    t.Alignment = Alignment.center;
                    t.AutoFit = AutoFit.Contents;
                    parcelPic.InsertTableAfterSelf(t).Alignment = Alignment.center;
                    //t.AutoFit = AutoFit.
                    //picHelper.insert(document, parcelPic, path.pathHelper.GetParcelpicPath(CompanyID, Parcelname), 592, 630);
                    parcelPic.AppendLine();
                }

                using (FontFamily fontFamily = new FontFamily("宋体"))
                {
                    var landusePic = document.InsertParagraph("国有土地使用证", false, format.formatHelper.SetParagraphFormat(fontFamily, 16, Color.Black));
                    //landusePic.Append("国有土地使用证").FontSize(16);
                    landusePic.AppendLine();
                    picHelper.insert(document, landusePic, PathManager.getSingleton().GetLandusepicPath(parcelID, false), 406, 570);
                }
                
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入总平面分布图模块
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        public void insertZPMFBT(DocX document, bool ischild, bool istrainings, NumoftitleHelper title)
        {
            try
            {
                Paragraph h1_2;
                Picture p = picHelper.getPic(document, PathManager.getSingleton().GetZPMFBpicPath( parcelID, false), 677, 769);
                if (p == null) return;
                if (ischild)
                {
                    title.Less3Zero();
                    h1_2 = document.InsertParagraph(title.num3title() + "总平面分布图");
                    h1_2.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_2.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else if (istrainings)
                {
                    title.Less4Zero();
                    h1_2 = document.InsertParagraph(title.num4title() + "总平面分布图");
                    h1_2.StyleName = "Heading4";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_2.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else
                {
                    title.Less2Zero();
                    h1_2 = document.InsertParagraph(title.num2title() + "总平面分布图");
                    h1_2.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_2.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                }
                h1_2.AppendLine();
                var Pic = document.InsertParagraph();
                Table t = tableHelper.parcelpicTabel(document, p, "总平面分布图");
                //picHelper.insert(document, Pic, path.pathHelper.GetZPMFBpicPath(CompanyID, Parcelname), 521, 567);
                Pic.InsertTableAfterSelf(t).Alignment = Alignment.center;
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                Pic.AppendLine();

                //var ownershipPic = document.InsertParagraph();
                //ownershipPic.AppendLine("房屋所有权证");
                //picHelper.insert(document, landusePic, Environment.CurrentDirectory + "\\公司\\国网江苏省电力公司高邮市供电公司\\" + Name + "\\房屋所有权证.jpg");//todo:图片路径要改活

            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入鸟瞰图模块
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        public void insertNKT(DocX document, bool ischild, bool istrainings, NumoftitleHelper title)
        {
            try
            {
                Paragraph h1_3;
                Picture p_Orth = picHelper.getPic(document, PathManager.getSingleton().GetAerialviewpicPath( parcelID, "正射", false), 308, 443);
                Picture p_front = picHelper.getPic(document, PathManager.getSingleton().GetAerialviewpicPath( parcelID, "前", false), 308, 443);
                Picture p_rear = picHelper.getPic(document, PathManager.getSingleton().GetAerialviewpicPath( parcelID, "后", false), 308, 443);
                Picture p_left = picHelper.getPic(document, PathManager.getSingleton().GetAerialviewpicPath( parcelID, "左", false), 308, 443);
                Picture p_right = picHelper.getPic(document, PathManager.getSingleton().GetAerialviewpicPath( parcelID, "右", false), 308, 443);
                if (p_Orth == null || p_front == null || p_rear == null || p_left == null || p_right == null) return;
                if (ischild)
                {
                    title.Less3Zero();
                    h1_3 = document.InsertParagraph(title.num3title() + "鸟瞰图（航拍）");
                    h1_3.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_3.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else if (istrainings)
                {
                    title.Less4Zero();
                    h1_3 = document.InsertParagraph(title.num4title() + "鸟瞰图（航拍）");
                    h1_3.StyleName = "Heading4";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_3.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else
                {
                    title.Less2Zero();
                    h1_3 = document.InsertParagraph(title.num2title() + "鸟瞰图（航拍）");
                    h1_3.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_3.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                }
                h1_3.AppendLine();
                var AerialViewPic = document.InsertParagraph();
                Table t = tableHelper.AerialviewpicTable(document, p_Orth, p_front, p_rear, p_left, p_right);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Contents;
                AerialViewPic.InsertTableAfterSelf(t).Alignment = Alignment.center;
                AerialViewPic.AppendLine();
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }
        }

        /// <summary>
        /// 插入分层分户平面图
        /// </summary>
        /// <param name="document"></param>
        /// <param name="lstFM"></param>
        /// <param name="p"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        public void insertFCFH(DocX document, List<FDXXtbl_country> lstFM, List<Parcelmodel> p, bool ischild, bool istrainings, NumoftitleHelper title)
        {
            Paragraph h1_4_1;

            try
            {
                Paragraph h1_4 = document.InsertParagraph("");

                FDXXtbl_country temp = new FDXXtbl_country();
                foreach (FDXXtbl_country fm in lstFM)
                {
                    if (temp != null && temp.ZDXX_MC == fm.ZDXX_MC && temp.FCXX_JZMC == fm.FCXX_JZMC)
                        continue;
                    temp = fm;

                    //if(fm.ZDXX_MC == Parcelname)
                    if (fm.ZDXX_ID == parcelID )
                    {


                        Picture p_deed = picHelper.getPic(document, PathManager.getSingleton().GetDeedPath( fm.FCXX_ID, false), 450, 886);
                        Picture p_virtual = picHelper.getPic(document, PathManager.getSingleton().GetVirtualmapPath( fm.FCXX_ID, false), 450, 886);
                        List<string> lstStr = txt.txtHelper.txtLines(PathManager.getSingleton().GetPlantxtPath( fm.FCXX_ID, false));
                        
                        if (h1_4.Text == "")
                        {
                            //if (p_deed != null || p_virtual != null || lstStr.Count != 0 || t != null)
                            //{
                                if (ischild)
                                {
                                    title.Less3Zero();
                                    h1_4 = document.InsertParagraph(title.num3title() + "分层分户平面图");
                                    h1_4.StyleName = "Heading3";
                                    using (FontFamily fontfamily = new FontFamily("宋体"))
                                    {
                                        h1_4.Color(Color.Black).FontSize(14).Font(fontfamily);
                                    }

                                }
                                else if (istrainings)
                                {
                                    title.Less4Zero();
                                    h1_4 = document.InsertParagraph(title.num4title() + "分层分户平面图");
                                    h1_4.StyleName = "Heading4";
                                    using (FontFamily fontfamily = new FontFamily("宋体"))
                                    {
                                        h1_4.Color(Color.Black).FontSize(14).Font(fontfamily);
                                    }
                                }
                                else
                                {
                                    title.Less2Zero();
                                    h1_4 = document.InsertParagraph(title.num2title() + "分层分户平面图");
                                    h1_4.StyleName = "Heading2";
                                    using (FontFamily fontfamily = new FontFamily("宋体"))
                                    {
                                        h1_4.Color(Color.Black).FontSize(16).Font(fontfamily);
                                    }
                                }
                            //}
                        }
                        if (ischild)
                        {
                            title.Less4Zero();
                            h1_4_1 = document.InsertParagraph(title.num4title() + fm.FCXX_JZMC + "分层分户平面图");
                            h1_4_1.StyleName = "Heading4";
                            using (FontFamily fontfamily = new FontFamily("宋体"))
                            {
                                h1_4_1.Color(Color.Black).FontSize(14).Font(fontfamily).Italic();
                            }

                        }
                        else if (istrainings)
                        {
                            //title.Less4Zero();
                            h1_4 = document.InsertParagraph(title.num5title() + "分层分户平面图");
                            h1_4.StyleName = "Heading5";
                            using (FontFamily fontfamily = new FontFamily("宋体"))
                            {
                                h1_4.Color(Color.Black).FontSize(14).Font(fontfamily);
                            }

                        }
                        else
                        {
                            title.Less3Zero();
                            h1_4_1 = document.InsertParagraph(title.num3title() + fm.FCXX_JZMC + "分层分户平面图");
                            h1_4_1.StyleName = "Heading3";
                            using (FontFamily fontfamily = new FontFamily("宋体"))
                            {
                                h1_4_1.Color(Color.Black).FontSize(14).Font(fontfamily);
                            }                           
                        }
                        //房产证
                        if (p_deed != null)
                        {
                            var tbl_deed = document.InsertParagraph();//房产证表格
                            Table t_deed = tableHelper.picTable(document, p_deed, "房产证");
                            t_deed.Alignment = Alignment.center;
                            t_deed.AutoFit = AutoFit.Contents;
                            tbl_deed.InsertTableAfterSelf(t_deed);
                        }
                        
                        //实景图
                        if (p_virtual != null)
                        {
                            var tbl_virtual = document.InsertParagraph();//实景图表格
                            Table t_virtual = tableHelper.picTable(document, p_virtual,"房产外墙面实景图");
                            t_virtual.Alignment = Alignment.center;
                            t_virtual.AutoFit = AutoFit.Contents;
                            tbl_virtual.InsertTableAfterSelf(t_virtual);
                        }
                        
                        //平面图
                        var title2 = document.InsertParagraph();//生产综合楼分层分户平面图标题
                        title2.Append("分层分户平面图").FontSize(16);
                        title2.Alignment = Alignment.center;

                        if(lstStr.Count > 0)
                        {
                            Table t = tableHelper.PlanTable(document, companyID, parcelID, fm.FCXX_ID, lstStr);
                            t.Alignment = Alignment.center;
                            t.AutoFit = AutoFit.Window;

                            document.InsertTable(t).Alignment = Alignment.center;
                        }                        
                        
                    }
                    

                    //if (fm.ZDXX_MC == Parcelname)
                    //{
                        
                    //}
                }
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }

        }

        /// <summary>
        /// 插入场地涉外管线布置图
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ischild">判断是否作为市级公司的下级</param>
        public void insertGXT(DocX document, bool ischild, bool istrainings, NumoftitleHelper title)
        {

            try
            {
                Paragraph h1_5;
                Picture p = picHelper.getPic(document, PathManager.getSingleton().GetPipelinepicPath( parcelID, false), 700, 619);
                if (p == null) return;
                if (ischild)
                {
                    title.Less3Zero();
                    h1_5 = document.InsertParagraph(title.num3title() + "场地涉外管线布置图");
                    h1_5.StyleName = "Heading3";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_5.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else if (istrainings)
                {
                    title.Less4Zero();
                    h1_5 = document.InsertParagraph(title.num4title() + "场地涉外管线布置图");
                    h1_5.StyleName = "Heading4";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_5.Color(Color.Black).FontSize(14).Font(fontfamily);
                    }
                }
                else
                {
                    title.Less2Zero();
                    h1_5 = document.InsertParagraph(title.num2title() + "场地涉外管线布置图");
                    h1_5.StyleName = "Heading2";
                    using (FontFamily fontfamily = new FontFamily("宋体"))
                    {
                        h1_5.Color(Color.Black).FontSize(16).Font(fontfamily);
                    }
                }
                var Pic = document.InsertParagraph();
                Table t = tableHelper.pipelinepicTable(document, p);
                t.Alignment = Alignment.center;
                t.AutoFit = AutoFit.Window;
                Pic.InsertTableAfterSelf(t).Alignment = Alignment.center;
                //picHelper.insert(document, Pic, path.pathHelper.GetPipelinepicPath(CompanyID, Parcelname), 536, 457);
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(parcelHelper), ex);
            }

        }

    }
}
