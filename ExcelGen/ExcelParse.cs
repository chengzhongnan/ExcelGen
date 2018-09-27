using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using System.Xml.Linq;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Reflection.Emit;
using System.Collections;

namespace ExcelTool
{
    /// <summary>
    /// Excel定义
    /// 1. Sheet名称（只能包含英文字符,下划线，数字）为表结构体名称（默认的Sheet1，Sheet2……会被忽略处理）
    /// 2. 每个Sheet第一列为结构体字段，必须为英文，如果为结构体，需要合并单元格
    /// 3. 第二列为结构体子字段，如果是普通字段，与第一列相同，如果时结构体，需要注明结构体每一个字段，如果是结构体数组，需要重复填写结构体字段
    /// 4. 第三列为第二列对应字段的数据类型，只支持基本数据类型，包括byte, short, int, long, double, float, bool, string
    /// 5. 第四列为字段中文描述，可以为任意字符，该文字描述将自动添加到生成的C#代码注释中
    /// 6. 第五行以下为数据，在留空的情况下byte, short, int, long, double, float默认值为0，bool默认值为false， string默认值为空字符串
    /// </summary>
    class ExcelParse
    {
        public ExcelParse(string fileName)
        {
            FileName = fileName;
            regInvalidSheetName = new Regex("^[Ss]heet\\d+$");
            regValidSheetName = new Regex("^[a-zA-Z1-9_]+$");
            Init();
        }

        private string FileName = string.Empty;
        private XSSFWorkbook workBook = null;
        private Regex regInvalidSheetName = null;
        private Regex regValidSheetName = null;
        public void Init()
        {
            try
            {
                using (FileStream file = new FileStream(FileName, FileMode.Open, FileAccess.Read))
                {
                    workBook = new XSSFWorkbook(file);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw e;
            }
        }

        public void DoParse(ref Dictionary<string, string> classDic, string xmlSavePath, string jsonSavePath)
        {
            var sheetNames = GetSheets();
            Dictionary<string, List<ExcelHeader>> headerDic = new Dictionary<string, List<ExcelHeader>>();
            foreach (var sheetName in sheetNames)
            {
                try
                {
                    var headers = GetHeader(sheetName);
                    headerDic[sheetName] = headers;

                    GenarateCSharpCode gc = new GenarateCSharpCode();
                    var classText = gc.GenarateClass(sheetName, headers);
                    classDic.Add(sheetName, classText);
                    var indexValues = SaveXMLData(sheetName, headers, 
                        xmlSavePath.Trim('\\') + "\\" + sheetName + ".xml",
                        jsonSavePath.Trim('\\') + "\\" + sheetName + ".json");
                    if (sheetName == "SystemConfig")
                    {
                        var enumText = gc.GenarateEnum(sheetName, indexValues);
                        classDic.Add(nameof(Enum) + sheetName, enumText);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Sheet [" + sheetName + "] header invalid");
                }
            }
        }

        /// <summary>
        /// 解析Excel文档
        /// </summary>
        public void DoParseTest(string outDllFileName)
        {
            var sheetNames = GetSheets();
            Dictionary<string, List<ExcelHeader>> headerDic = new Dictionary<string, List<ExcelHeader>>();
            Dictionary<string, string> classDic = new Dictionary<string, string>();
            foreach(var sheetName in sheetNames)
            {
                try
                {
                    var headers = GetHeader(sheetName);
                    headerDic[sheetName] = headers;

                    GenarateCSharpCode gc = new GenarateCSharpCode();
                    var classText = gc.GenarateClass(sheetName, headers);
                    classDic.Add(sheetName, classText);
                    var indexValues = SaveXMLData(sheetName, headers, sheetName + ".xml", sheetName + ".json");

                    if(sheetName == "SystemConfig")
                    {
                        var enumText = gc.GenarateEnum(sheetName, indexValues);
                        classDic.Add(nameof(Enum) + sheetName, enumText);
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Sheet [" + sheetName + "] header invalid");
                }
            }

            var assemblyText = Resource1.ClassTemplete;
            assemblyText += "\r\n{\r\n";
            assemblyText += Resource1.ScriptBase;
            assemblyText += "\r\n\r\n";

            foreach(var classTextkv in classDic)
            {
                assemblyText += classTextkv.Value;
                assemblyText += "\r\n";
            }
            
            assemblyText += "\r\n}\r\n";

            DynamicCompile compile = new DynamicCompile();
            var asm = compile.Compile(assemblyText, outDllFileName);

            //var type = asm.GetType("ExcelTables." + sheetName);
            //SaveXMLData(sheetName, headers, type, sheetName + ".xml");
        }

        public List<string> SaveXMLData(string sheetName, List<ExcelHeader> headers, string fileName, string jsonfileName)
        {
            var sheet = workBook.GetSheet(sheetName);
            XElement xRoot = new XElement("Root");
            List<string> enumIndex = new List<string>();
            for(var i = 4; i < sheet.LastRowNum + 1; i++)
            {
                var rowIndexValue = string.Empty;
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    // 跳过空行
                    break;
                }
                var xRow = GenarateRowData(sheetName, headers, row, ref rowIndexValue);
                if (string.IsNullOrEmpty(rowIndexValue))
                {
                    // 跳过索引为空的行
                    break;
                }
                xRoot.Add(xRow);
                enumIndex.Add(rowIndexValue);
            }
            xRoot.Save(fileName);
            SaveJson(xRoot, jsonfileName);
            return enumIndex;
        }

        public string GetCellValue(ICell cell)
        {
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        return string.Empty;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Formula:
                        {
                            if (cell.CachedFormulaResultType == CellType.String)
                            {
                                return cell.StringCellValue;
                            }
                            if (cell.CachedFormulaResultType == CellType.Numeric)
                            {
                                return cell.NumericCellValue.ToString();
                            }
                            if (cell.CachedFormulaResultType == CellType.Boolean)
                            {
                                return cell.BooleanCellValue.ToString();
                            }
                            return string.Empty;
                        }
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Unknown:
                        return string.Empty;
                }
                return string.Empty;
            }
            catch(Exception ex)
            {
                //Console.WriteLine(ex);
                return string.Empty;
            }
        }

        public XElement GenarateRowData(string eleName , List<ExcelHeader> headers, IRow row, ref string rowEnumText)
        {
            XElement xEle = new XElement(eleName);
            foreach(var header in headers)
            {
                XElement xEleValue = new XElement(header.Name);
                if(header.ColumnStart == header.ColumnEnd)
                {
                    var cell = row.GetCell(header.ColumnStart);
                    xEleValue.Value = GetCellValue(cell);
                }
                else
                {
                    var mutiXml = GetMutiHeaderXml(header, row);
                    foreach(var subMutiXml in mutiXml)
                    {
                        xEleValue.Add(subMutiXml);
                    }
                }

                if(header.ColumnStart == 0)
                {
                    rowEnumText = xEleValue.Value;
                }

                xEle.Add(xEleValue);
            }
            return xEle;
        }

        public List<XElement> GetMutiHeaderXml(ExcelHeader header, IRow row)
        {
            var xEleAll = new List<XElement>();
            if (header.SubNode == null)
                return xEleAll;

            foreach(var subHeaderDic in header.SubNode)
            {
                if(subHeaderDic.Value.Count == 1)
                {
                    try
                    {
                        for (int i = header.ColumnStart; i <= header.ColumnEnd; i++)
                        {
                            var cell = row.Cells.Find(x => x.ColumnIndex == i);
                            if (cell == null || cell.CellType == CellType.Blank)
                            {
                                continue;
                            }
                            var xSub = new XElement(header.Name + "Element", GetCellValue(cell));
                            xEleAll.Add(xSub);
                        }
                        break;
                    }
                    catch(Exception ex)
                    {
                        throw new Exception("[" + row.Sheet.SheetName + "][" + header.Name + "] invalid", ex);
                    }
                }
                else
                {
                    try
                    {
                        var xSub = new XElement(header.Name + "Element");

                        bool bAdd = true;
                        bool bFirst = true;
                        foreach (var sh in subHeaderDic.Value)
                        {
                            var col = sh.ColumnStart;
                            var cell = row.Cells.Find(x => x.ColumnIndex == col);
                            if (bFirst &&( cell == null || cell.CellType == CellType.Blank))
                            {
                                bAdd = false;
                                break;
                            }
                            bFirst = false;
                            var xSubEle = new XElement(sh.Name, GetCellValue(cell));
                            xSub.Add(xSubEle);
                        }
                        if (bAdd)
                        {
                            xEleAll.Add(xSub);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("[" + row.Sheet.SheetName + "][" + header.Name + "] invalid", ex);
                    }
                }
            }
            return xEleAll;
        }

        string[] GetSheets()
        {
            List<string> sheetNames = new List<string>();
            for(ushort i = 0; i < workBook.Count; i++)
            {
                var name = workBook.GetSheetName(i);
                if(string.IsNullOrEmpty(name))
                {
                    break;
                }
                if(regInvalidSheetName.IsMatch(name))
                {
                    // 跳过没起名字的Sheet页面
                    continue;
                }
                if (regValidSheetName.IsMatch(name))
                {
                    sheetNames.Add(name);
                }
            }
            return sheetNames.ToArray();
        }

        /// <summary>
        /// 取得页面表头信息
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        List<ExcelHeader> GetHeader(string sheetName)
        {
            List<ExcelHeader> headerList = new List<ExcelHeader>();
            ISheet sheet = workBook.GetSheet(sheetName);

            var row_Main = sheet.GetRow(0);
            var row_Detail = sheet.GetRow(1);
            var row_Type = sheet.GetRow(2);
            var row_Desc = sheet.GetRow(3);
            for(int i = 0; i < row_Main.Cells.Count; )
            {
                var cell = row_Main.Cells[i];
                if(cell.IsMergedCell)
                {
                    var mergedCells = GetMergedCells(row_Main, cell.ColumnIndex);
                    var header = GetMutiExcelHeader(row_Main, row_Detail, row_Type, row_Desc, mergedCells);
                    headerList.Add(header);
                    i += mergedCells.Count;
                }
                else
                {
                    var header = GetSingleExcelHeader(row_Main, row_Type, row_Desc, cell.ColumnIndex);
                    headerList.Add(header);
                    i++;
                }
            }

            return headerList;
        }

        /// <summary>
        /// 取得合并的单元格序号
        /// </summary>
        /// <param name="row"></param>
        /// <param name="nStart"></param>
        /// <returns></returns>
        List<int> GetMergedCells(IRow row, int nStart)
        {
            List<int> colIndex = new List<int>() { nStart };
            for(var i = nStart + 1; i < row.LastCellNum; i++)
            {
                var cell = row.GetCell(i);
                if (string.IsNullOrEmpty( cell.RichStringCellValue.ToString()))
                {
                    colIndex.Add(i);
                }
                else
                {
                    break;
                }
            }
            return colIndex;
        }

        /// <summary>
        /// 取得某一列的表头信息
        /// </summary>
        /// <param name="rowMain"></param>
        /// <param name="rowType"></param>
        /// <param name="rowDesc"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        ExcelHeader GetSingleExcelHeader(IRow rowMain, IRow rowType, IRow rowDesc, int col)
        {
            ExcelHeader header = new ExcelHeader();
            header.Name = rowMain.GetCell(col).StringCellValue;
            header.Type = rowType.GetCell(col).StringCellValue;
            header.Describe = rowDesc.GetCell(col).StringCellValue;
            header.ColumnStart = col;
            header.ColumnEnd = col;
            return header;
        }

        /// <summary>
        /// 取得多列表头信息
        /// </summary>
        /// <param name="rowMain"></param>
        /// <param name="rowType"></param>
        /// <param name="rowDesc"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        ExcelHeader GetMutiExcelHeader(IRow rowMain, IRow rowDetail,  IRow rowType, IRow rowDesc, List<int> cols)
        {
            ExcelHeader header = new ExcelHeader();
            header.ColumnStart = cols.Min();
            header.ColumnEnd = cols.Max();
            header.Describe = string.Empty;
            header.Name = rowMain.GetCell(header.ColumnStart).StringCellValue;
            header.SubNode = new Dictionary<int, List<ExcelHeader>>();

            int structLength = GetMutiExcelStructLength(rowDetail, cols);
            header.SubClassFieldLength = structLength;
            if(structLength == 1)
            {
                header.Type = "List<" + rowType.GetCell(header.ColumnStart).StringCellValue + ">";
            }
            else
            {
                header.Type = "List<" + header.Name + ">";
            }

            for (var i = header.ColumnStart; i <= header.ColumnEnd; i += structLength)
            {
                header.SubNode[header.SubNode.Count] = GetStructHeader(rowDetail, rowType, rowDesc, i, structLength);
            }

            return header;
        }

        /// <summary>
        /// 取得结构体长度
        /// </summary>
        /// <param name="rowDetail"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        int GetMutiExcelStructLength(IRow rowDetail, List<int> cols)
        {
            string firstName = string.Empty;
            int nCycle = cols.Count;
            for (var i = cols.Min(); i <= cols.Max(); i ++)
            {
                var name = rowDetail.GetCell(i).StringCellValue;
                if(firstName == string.Empty)
                {
                    firstName = name;
                }
                else if(firstName == name)
                {
                    nCycle = i - cols.Min();
                    break;
                }
            }

            return nCycle;
        }

        /// <summary>
        /// 取得结构体的ExcelHeader
        /// </summary>
        /// <param name="rowDetail"></param>
        /// <param name="rowType"></param>
        /// <param name="rowDesc"></param>
        /// <param name="nStart"></param>
        /// <param name="nLen"></param>
        /// <returns></returns>
        List<ExcelHeader> GetStructHeader(IRow rowDetail, IRow rowType, IRow rowDesc, int nStart, int nLen)
        {
            List<ExcelHeader> headerList = new List<ExcelHeader>();
            for(int i = nStart; i < nStart + nLen; i++)
            {
                ExcelHeader header = new ExcelHeader();
                header.ColumnStart = i;
                header.ColumnEnd = i;
                header.Name = rowDetail.GetCell(i).StringCellValue;
                header.Type = rowType.GetCell(i).StringCellValue;
                header.Describe = rowDesc.GetCell(i).StringCellValue;

                headerList.Add(header);
            }
            return headerList;
        }

        private void SaveJson(XElement xml, string fileName)
        {
            ArrayList jsonObject = new ArrayList();
            foreach (var ele in xml.Elements())
            {
                Dictionary<string, object> objConfig = new Dictionary<string, object>();
                foreach (var col in ele.Elements())
                {
                    if (col.Elements().Count() > 0)
                    {
                        List<object> subObjectList = new List<object>();
                        foreach (var subCol in col.Elements())
                        {
                            if (subCol.Elements().Count() > 0)
                            {
                                Dictionary<string, object> objSub = new Dictionary<string, object>();
                                foreach (var subColEle in subCol.Elements())
                                {
                                    objSub[subColEle.Name.LocalName] = subColEle.Value;
                                }
                                subObjectList.Add(objSub);
                            }
                            else
                            {
                                subObjectList.Add(subCol.Value);
                            }
                        }
                        objConfig[col.Name.LocalName] = subObjectList;
                    }
                    else
                    {
                        objConfig[col.Name.LocalName] = col.Value;
                    }
                }
                jsonObject.Add(objConfig);
            }

            var jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObject);

            File.WriteAllText(fileName, jsonStr);
        }
    }

    /// <summary>
    /// 编译目录下所有Excel
    /// </summary>
    public class ExcelParseFolder
    {
        static ExcelParseFolder _Instance = null;
        public static ExcelParseFolder Instance => _Instance ?? (_Instance = new ExcelParseFolder());

        public void DoParseFolder(string folderPath, string dllSavePath, string xmlSavePath, string jsonSavePath)
        {
            Dictionary<string, string> classDic = new Dictionary<string, string>();
            foreach(var file in Directory.GetFiles(folderPath))
            {
                if(file.EndsWith(".xlsx"))
                {
                    ExcelParse parser = new ExcelParse(file);
                    parser.DoParse(ref classDic, xmlSavePath, jsonSavePath);
                }
            }

            var assemblyText = Resource1.ClassTemplete;
            assemblyText += "\r\n{\r\n";
            assemblyText += Resource1.ScriptBase;
            assemblyText += "\r\n\r\n";

            foreach (var classTextkv in classDic)
            {
                assemblyText += classTextkv.Value;
                assemblyText += "\r\n";
            }

            assemblyText += "\r\n}\r\n";

            DynamicCompile compile = new DynamicCompile();
            compile.Compile(assemblyText, dllSavePath + "\\TableConfig.dll");
        }
    }

    /// <summary>
    /// Excel文件头描述
    /// </summary>
    public class ExcelHeader
    {
        /// <summary>
        /// 字段英文名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 字段类型，如果是结构体，那么类型字符串是 List<Name>，其他的是基本类型
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// 子类型，如果Type是基本字符串，那么该字段是空值，否则是所有子类型
        /// </summary>
        public Dictionary<int, List<ExcelHeader>> SubNode { get; set; }

        /// <summary>
        /// 字段描述，将写入到注释中
        /// </summary>
        public string Describe { get; set; }

        /// <summary>
        /// 起始列
        /// </summary>
        public int ColumnStart { get; set; }
        /// <summary>
        /// 终止列
        /// </summary>
        public int ColumnEnd { get; set; }

        public int SubClassFieldLength { get; set; }
    }
}
