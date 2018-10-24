using Newtonsoft.Json;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelTool
{
    class Program
    {
        static void Main(string[] args)
        {
            if(!System.IO.Directory.Exists("./dll"))
            {
                System.IO.Directory.CreateDirectory("./dll");
            }
            if(!System.IO.Directory.Exists("./xml"))
            {
                System.IO.Directory.CreateDirectory("./xml");
            }
            if(!System.IO.Directory.Exists("./json"))
            {
                System.IO.Directory.CreateDirectory("./json");
            }
            ExcelParseFolder.Instance.DoParseFolder("./xlsx", "./dll", "./xml", "./json");

            Md5HashFiles("./json");
            //var startTick = Environment.TickCount;
            //for (var o = 0; o < 1000; o++)
            //{
            //    MyDPTable.Instance.LoadTable("./xml/DP_Table.xml");
            //}
            //Console.WriteLine(Environment.TickCount - startTick);
            
            // AutoLoadConfig.Instance.LoadConfig("./xml");
            //ExcelTables.AutoLoadConfig.Instance.RegistryAssembly(typeof(Program).Assembly);
            //ExcelTables.AutoLoadConfig.Instance.LoadConfig("./xml");
        }

        class FileHash
        {
            [JsonProperty("file")]
            public string File { get; set; }
            [JsonProperty("md5")]
            public string MD5 { get; set; }
            [JsonProperty("size")]
            public long Size { get; set; }
        }

        static void Md5HashFiles(string folder)
        {
            try
            {
                if (!folder.EndsWith("/"))
                {
                    folder += "/";
                }

                var configFile = System.IO.File.ReadAllLines("./md5.txt");
                var hashFiles = new System.Collections.Generic.List<FileHash>();
                foreach (var file in configFile)
                {
                    if (System.IO.File.Exists(folder + file))
                    {
                        FileHash fh = new FileHash() { File = file };
                        fh.MD5 = GetFileMD5(folder + file);
                        fh.Size = File.ReadAllBytes(folder + file).Length;
                        hashFiles.Add(fh);
                    }
                }

                System.IO.File.WriteAllText(folder + "md5.json", JsonConvert.SerializeObject(hashFiles));
            }
            catch (System.Exception)
            {

            }
        }

        static string GetFileMD5(string fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Open);
            System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            byte[] retVal = md5.ComputeHash(file);
            file.Close();

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < retVal.Length; i++)
            {
                sb.Append(retVal[i].ToString("x2"));
            }
            return sb.ToString();
        }
    }

    //public class DP_Table : ScriptTableBase<DP_Table, int>
    //{
    //    static DP_Table _Instance = null;
    //    public static new DP_Table Instance
    //    {
    //        get
    //        {
    //            if(_Instance == null)
    //            {
    //                _Instance = new DP_Table();
    //            }
    //            return _Instance;
    //        }
    //    }
    //    // Properties
    //    public int ActivityId { get; set; }
    //    public string Describe { get; set; }
    //    public int DPParam1 { get; set; }
    //    public int DPParam2 { get; set; }
    //    public int DPParam3 { get; set; }
    //    public int DPPoint { get; set; }
    //    public int DPSubType { get; set; }
    //    public int DPType { get; set; }
    //    [ExcelIndex]
    //    public int Index { get; set; }
    //    public int TargetValue { get; set; }

    //    public virtual string XMLFile
    //    {
    //        get
    //        {
    //            return "DP_Table.xml";
    //        }
    //    }
    //}

    //public class MyDPTable : DP_Table
    //{
    //    static MyDPTable _Instance = null;
    //    public static new MyDPTable Instance => _Instance ?? (_Instance = new MyDPTable());

    //    Dictionary<int, List<DP_Table>> _SubTypeTables = new Dictionary<int, List<DP_Table>>();
    //    public Dictionary<int, List<DP_Table>> SubTypeTables => _SubTypeTables;
    //    public override void Add(DP_Table table)
    //    {
    //        base.Add(table);

    //        if(!SubTypeTables.ContainsKey(table.DPSubType))
    //        {
    //            SubTypeTables[table.DPSubType] = new List<DP_Table>();
    //        }
    //        SubTypeTables[table.DPSubType].Add(table);
    //    }
    //    public override void LoadTable(string fileName)
    //    {
    //        Tables.Clear();
    //        SubTypeTables.Clear();
    //        base.LoadTable(fileName);
    //    }
    //}
}
