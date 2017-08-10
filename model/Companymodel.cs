using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.model
{
    public class Companymodel
    {
        #region 成员变量
        public string name { get; set; }//公司名称
        public decimal ID { get; set; }//公司ID
        public string property { get; set; }//公司性质：供电公司、本部、直属单位等等。
        #endregion
    }
}
