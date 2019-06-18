using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    class GenarateCSharpCode : GenarateCodeBase
    {
        public override string GenarateEnum(string name, List<string> enumNames)
        {
            string enumName = nameof(Enum) + name;
            StringBuilder sb = new StringBuilder();
            sb.Append("public enum ");
            sb.Append(enumName);
            sb.Append("\r\n");
            sb.Append("{\r\n");
            foreach(var eName in enumNames)
            {
                sb.Append(eName);
                sb.Append(",\r\n");
            }
            sb.Append("}\r\n");
            return sb.ToString();
        }

        public override string GenarateClass(string name, List<ExcelHeader> excelHeader)
        {
            string type = GetExcelIndexType(excelHeader);
            StringBuilder sb = new StringBuilder();
            sb.Append("public class ");
            sb.Append(name);
            sb.Append(" : ScriptTableBase<");
            sb.Append(name);
            sb.Append(",");
            sb.Append(type);
            sb.Append(">\r\n");
            sb.Append("{\r\n");
            GenarateInstance(ref sb, name);
            GenarateXMLFileName(ref sb, name);
            foreach (var eh in excelHeader)
            {
                GenarateProperty(ref sb, eh);
            }

            GenarateTGet(ref sb, name, type);
            GenarateTClear(ref sb, name);

            sb.Append("\r\n}");
            return sb.ToString();
        }
        
        private void GenarateTGet(ref StringBuilder sb, string name, string typeName)
        {
            sb.Append("\tpublic static ");
            sb.Append(name);
            sb.Append(" TGet(");
            sb.Append(typeName);
            sb.Append(" Index)\r\n");
            sb.Append("\t{");
            sb.Append("\t\treturn Instance.Get(Index);\r\n");
            sb.Append("\t}\r\n\r\n");
        }

        private void GenarateTClear(ref StringBuilder sb, string name)
        {
            sb.Append("\tpublic static void TClear()\r\n{\r\nInstance.Clear();\r\n}\r\n\r\n");
        }

        private void GenarateInstance(ref StringBuilder sb, string name)
        {
            //static DP_Table _Instance = null;
            //public static new DP_Table Instance
            //{
            //    get
            //    {
            //        if (_Instance == null)
            //        {
            //            _Instance = new DP_Table();
            //        }
            //        return _Instance;
            //    }
            //}
            sb.Append("\tstatic ");
            sb.Append(name);
            sb.Append(" _Instance = null;\r\n");
            sb.Append("\tpublic static new ");
            sb.Append(name);
            sb.Append(" Instance\r\n");
            sb.Append("\t{\r\n");
            sb.Append("\t\tget\r\n");
            sb.Append("\t\t{\r\n");
            sb.Append("\t\t\tif (_Instance == null)\r\n");
            sb.Append("\t\t\t{\r\n");
            sb.Append("\t\t\t\t_Instance = new ");
            sb.Append(name);
            sb.Append("();\r\n");
            sb.Append("\t\t\t}\r\n");
            sb.Append("return _Instance;\r\n");
            sb.Append("\t\t}\r\n");
            sb.Append("\t}\r\n");
            sb.AppendLine();
        }

        private void GenarateXMLFileName(ref StringBuilder sb, string name)
        {
            //public virtual string XMLFile
            //{
            //    get
            //    {
            //        return "DP_Table.xml";
            //    }
            //}
            sb.Append("\tpublic virtual string XMLFile\r\n");
            sb.Append("\t{\r\n");
            sb.Append("\t\tget\r\n");
            sb.Append("\t\t{\r\n");
            sb.Append("\t\t\treturn \"");
            sb.Append(name);
            sb.Append(".xml\";\r\n");
            sb.Append("\t\t}\r\n");
            sb.Append("\t}\r\n");
            sb.AppendLine();
        }

        private string GetExcelIndexType(List<ExcelHeader> excelHeader)
        {
            foreach(var eh in excelHeader)
            {
                if(eh.ColumnStart == 0 && eh.ColumnEnd == 0)
                {
                    return eh.Type;
                }
            }
            return "int";
        }

        void GenaratePropertyWithType(ref StringBuilder sb, ExcelHeader eh, string type, string name)
        {
            if(eh.ColumnStart == 0)
            {
                sb.Append("\t[ExcelIndex]\r\n");
            }
            sb.Append("\t/// <summary>\r\n\t/// ");
            sb.Append(eh.Describe.Replace("\n","").Replace("\r",""));
            sb.Append("\r\n\t/// </summary>\r\n");
            sb.Append("\tpublic ");
            sb.Append(type);
            sb.Append(" ");
            sb.Append(name);
            sb.Append(" {get; set;}\r\n");
        }

        void GenarateProperty(ref StringBuilder sb, ExcelHeader eh)
        {
            if(eh.ColumnStart == eh.ColumnEnd || eh.SubClassFieldLength == 1)
            {
                GenaratePropertyWithType(ref sb, eh, eh.Type, eh.Name);
            }
            else if (eh.SubNode.Count == 1)
            {
                GenarateSubClass(ref sb, eh.Type, eh.SubNode[0]);
                GenaratePropertyWithType(ref sb, eh, eh.Type, eh.Name);
            }
            else
            {
                GenarateSubClass(ref sb, eh.Name ,eh.SubNode[0]);
                GenaratePropertyWithType(ref sb, eh, eh.Type, eh.Name + "List");
            }
        }

        public void GenarateSubClass(ref StringBuilder sb, string name ,List<ExcelHeader> subHeader)
        {
            sb.Append("\tpublic class ");
            sb.Append(name);
            sb.Append("\r\n\t{\r\n");
            foreach (var eh in subHeader)
            {
                sb.Append("\t\t/// <summary>\r\n\t\t/// ");
                sb.Append(eh.Describe);
                sb.Append("\r\n\t\t/// </summary>\r\n");
                sb.Append("\t\tpublic ");
                sb.Append(eh.Type);
                sb.Append(" ");
                sb.Append(eh.Name);
                sb.Append(" {get; set;}\r\n");
            }
            sb.Append("\r\n\t}\r\n");
        }

        public void GenarateXmlData()
        {

        }

        public void GenarateLoadFunction(ref StringBuilder sb, List<ExcelHeader> excelHeader)
        {
            sb.Append("\tpublic void Load(string fileName)\r\n");
            sb.Append("\t{\r\n");
            sb.Append("\t\tConsole.WriteLine(\"Load \"+fileName);");
            sb.Append("\t}\r\n");
        }

        public void GenarateSaveFunction(ref StringBuilder sb, List<ExcelHeader> excelHeader)
        {
            sb.Append("\tpublic void Save(string fileName)\r\n");
            sb.Append("\t{\r\n");
            sb.Append("\t\tConsole.WriteLine(\"Save \"+fileName);");
            sb.Append("\t}\r\n");
        }
    }
}
