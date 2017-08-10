using CreateWord.createcompanies;
using CreateWord.DB;
using CreateWord.log;
using CreateWord.model;
using CreateWord.parcels;
using CreateWord.table;
using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord
{
    class Program
    {
        static string CompanyName = "";
        static string CompanyID = "";
        //static int companylevel = 0;//公司级别。1是省级公司，2是市级公司，3是县级公司

        static void Main(string[] args)
        {
            try
            {
                //试验数据：args[0]是省公司，args[1]是扬州市公司，args[2]是高邮市公司
                CompanyID = args[0].ToString();
                Program.show();
                //工厂方法模式
                chooseFactory(Int32.Parse(CompanyID));

                Console.Read();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(Program), ex);
            }
        }

        /// <summary>
        /// 为了能够在静态的Main函数中调用非静态的createword函数
        /// </summary>
        public static void show()
        {
            try
            {               
                new Program().getName(Int32.Parse(CompanyID));
            }
            catch (System.Exception ex)
            {
                LogHelper.WriteLog(typeof(Program), ex);
            }
        }

        /// <summary>
        /// 公司ID转换成名称
        /// </summary>
        /// <param name="id"></param>
        void getName(int id)
        {
            CompanyName = DBhelper.GetCompanyName(id);
        }

        static void chooseFactory(int id)
        {
            int level = judgeLevel(id);
            if (level == 1) //省级公司
            {
                IFactory cpyFactory = new ProvinceFactory();
                CreateCompany createCompany = cpyFactory.createcompany();
                createCompany.CompanyID = id;
                createCompany.CompanyName = CompanyName;
                createCompany.Childcompany = DBhelper.GetChildcompany(id);
                createCompany.createword();
            }
            else if (level == 2) //市级公司
            {
                IFactory cpyFactory = new CityFactory();
                CreateCompany createCompany = cpyFactory.createcompany();
                createCompany.CompanyID = id;
                createCompany.CompanyName = CompanyName;
                createCompany.Childcompany = DBhelper.GetChildcompany(id);
                createCompany.createword();
            }
            else if (level == 3) //县级公司
            {
                IFactory cpyFactory = new CountryFactory();
                CreateCompany createCompany = cpyFactory.createcompany();
                createCompany.CompanyID = id;
                createCompany.CompanyName = CompanyName;
                createCompany.Childcompany = DBhelper.GetChildcompany(id);
                createCompany.createword();
            }
        }

        /// <summary>
        /// 判断公司级别。1是省级公司，2是市级公司，3是县级公司。
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        static int judgeLevel(int id)
        {
            int i = 0;
            if (id < 10) i = 1;//省级公司
            else if (id >= 1000 && id < 10000) i = 2;//市级公司
            else i = 3;//县级公司
            return i;
        }
    }
}
