using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateWord.title
{
    public class NumoftitleHelper
    {
        private int num1;//一级标题
        private int num2;//二级标题
        private int num3;//三级标题
        private int num4;//四级标题
        private int num5;//五级标题
        public int Num1
        {
            get { return num1; }
            set { num1 = value; }
        }
        public int Num2
        {
            get { return num2; }
            set { num2 = value; }
        }
        public int Num3
        {
            get { return num3; }
            set { num3 = value; }
        }
        public int Num4
        {
            get { return num4; }
            set { num4 = value; }
        }
        public int Num5
        {
            get { return num5; }
            set { num5 = value; }
        }
        public NumoftitleHelper()
        {
            num1 = 0;
            num2 = 0;
            num3 = 0;
            num4 = 0;
            num5 = 0;
        }
        #region 标题序号生成
        /// <summary>
        /// 一级标题序号生成
        /// </summary>
        /// <returns></returns>
        public string num1title()
        {
            num1++;
            string title = num1.ToString() + ".  ";
            return title;
        }
        /// <summary>
        /// 二级标题序号生成
        /// </summary>
        /// <returns></returns>
        public string num2title()
        {
            num2++;
            string title = num1.ToString() + "." + num2.ToString() + "  ";
            return title;
        }
        /// <summary>
        /// 三级标题序号生成
        /// </summary>
        /// <returns></returns>
        public string num3title()
        {
            num3++;
            string title = num1.ToString() + "." + num2.ToString() + "." + num3.ToString() + "  ";
            return title;
        }
        /// <summary>
        /// 四级标题序号生成
        /// </summary>
        /// <returns></returns>
        public string num4title()
        {
            num4++;
            string title = num1.ToString() + "." + num2.ToString() + "." + num3.ToString() + "." + num4.ToString() + "  ";
            return title;
        }
        /// <summary>
        /// 五级标题序号生成
        /// </summary>
        /// <returns></returns>
        public string num5title()
        {
            num5++;
            string title = num1.ToString() + "." + num2.ToString() + "." + num3.ToString() + "." + num4.ToString() + "." + num5.ToString() + "  ";
            return title;
        }
        #endregion

        #region 重置为零

        /// <summary>
        /// 一级标题以下全置零
        /// </summary>
        public void Less1Zero()
        {
            num2 = 0;
            num3 = 0;
            num4 = 0;
            num5 = 0;
        }
        /// <summary>
        /// 二级标题以下全置零
        /// </summary>
        public void Less2Zero()
        {
            num3 = 0;
            num4 = 0;
            num5 = 0;
        }
        /// <summary>
        /// 三级标题以下全置零
        /// </summary>
        public void Less3Zero()
        {
            num4 = 0;
            num5 = 0;
        }
        /// <summary>
        /// 四级标题以下全置零
        /// </summary>
        public void Less4Zero()
        {
            num5 = 0;
        }

        #endregion


    }
}
