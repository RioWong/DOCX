using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.model
{
    public class Parcelmodel
    {
        #region 成员变量
        public int id { get; set; } //宗地ID
        public string name { get; set; } //宗地名称
        public int num { get; set; }//宗地里面建筑的个数。（表格里面合并单元格时需要）
        #endregion
    }

    public class Buildingmodel
    {
        #region 成员变量
        public int id { get; set; } //建筑ID
        public string name { get; set; } //建筑名称
        public int num { get; set; }//建筑里使用功能的个数。（表格里面合并单元格时需要）
        #endregion
    }
}
