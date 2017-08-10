using CreateWord.createcompanies;
using CreateWord.DB;
using CreateWord.listener;
using CreateWord.log;
using RealEstate.Logic;
using System;
using System.IO;

namespace CreateWord
{
    public class CREATE
    {
        log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private string CompanyName = "";
        private int CompanyID = -1;
        private string WordPath = "";//文档路径
        private string WordName = "";//文档名称
        private IDocCompilationListener docCompilationListener = null;
        //static int companylevel = 0;//公司级别。1是省级公司，2是市级公司，3是县级公司

        public CREATE( IDocCompilationListener docCompilationListener)
        {
            this.docCompilationListener = docCompilationListener;
        }

        /// <summary>
        /// 通过公司ID生成汇编文档
        /// </summary>
        /// <param name="companyID">公司ID</param>
        public void compilationdocument(int companyID)
        {
            //试验数据：args[0]是省公司，args[1]是扬州市公司，args[2]是高邮市公司
            CompanyID = companyID;
            WordName = Guid.NewGuid().ToString() + ".docx";  //文档名称采用生成的随机名称
            WordPath = Path.Combine(PathManager.getSingleton().GetBasePath( PathType.todoRootPath ), WordName);//生成的汇编文档保存在临时目录
            if (CompanyID <= 0)
            {
                string msg = "传入的公司ID有问题，请检查: CompanyID = " + CompanyID.ToString();
                LogHelper.WriteLog(typeof(CREATE), "传入的公司ID有问题，请检查！");
                if (docCompilationListener != null)
                {
                    docCompilationListener.DocCompleted(new DocCompilationArg(
                        CompanyID, WordPath, DocCompilationStatus.Fail, msg));
                }
                throw new Exception(msg);
            }

            CompanyName = DBhelper.GetCompanyName(CompanyID);
            log.InfoFormat("开始生成文档：{0}", CompanyName);

            //工厂方法模式
            createWord(CompanyID);
        }

        private void createWord(int id)
        {
            //int level = judgeLevel(id);
            int level = CompanyManager.getSingleton().GetCompanyLevel(id) + 1;
            CreateCompany createCompany = null;
            if (level == 1) //省级公司
            {
                createCompany = DocFactory.createcompany(DocCompanyType.Provice, docCompilationListener);
            }
            else if (level == 2) //市级公司
            {
                createCompany = DocFactory.createcompany(DocCompanyType.City, docCompilationListener);
            }
            else if (level == 3) //县级公司
            {
                createCompany = DocFactory.createcompany(DocCompanyType.Country, docCompilationListener);
            }

            if (createCompany == null) return;

            createCompany.CompanyID = id;
            createCompany.CompanyName = CompanyName;
            createCompany.Childcompany = DBhelper.GetChildcompany(id);
            createCompany.createword(WordPath);
        }

        ///// <summary>
        ///// 判断公司级别。1是省级公司，2是市级公司，3是县级公司。
        ///// </summary>
        ///// <param name="id"></param>
        ///// <returns></returns>
        //private static int judgeLevel(int id)
        //{
        //    int i = 0;
        //    if (id < 10) i = 1;//省级公司
        //    else if (id >= 1000 && id < 10000) i = 2;//市级公司
        //    else i = 3;//县级公司
        //    return i;
        //}
    }
}

