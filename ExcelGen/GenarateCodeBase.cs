using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    abstract class GenarateCodeBase
    {
        /// <summary>
        /// 生成类代码
        /// </summary>
        /// <param name="name">类名称</param>
        /// <param name="excelHeader">类具体数据</param>
        /// <returns></returns>
        public abstract string GenarateClass(string name, List<ExcelHeader> excelHeader);

        public virtual string GenarateEnum(string name, List<string> enumNames)
        {
            return string.Empty;
        }
    }
}
