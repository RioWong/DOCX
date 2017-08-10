
using CreateWord.listener;
using System;
namespace CreateWord.createcompanies
{
    public enum DocCompanyType { Provice, City, Country }

    /// <summary>
    /// 专门负责生产“公司文档”的工厂
    /// </summary>
    class DocFactory 
    {
        public static CreateCompany createcompany( DocCompanyType docCompanyType,IDocCompilationListener docCompilationListener )
        {
            switch( docCompanyType )
            {
                case DocCompanyType.Provice:
                    return new CreateProvinceCompany(docCompilationListener);
                case DocCompanyType.City:
                    return new CreateCityCompany(docCompilationListener);
                case DocCompanyType.Country:
                    return new CreateCountryCompany(docCompilationListener);
            }
            throw new Exception("非法文档类型");
        }
    }

}
