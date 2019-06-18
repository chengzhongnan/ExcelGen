using ExcelTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelTool.AutoLoadConfig.Instance.RegistryAssembly(Assembly.GetEntryAssembly());
            ExcelTool.AutoLoadConfig.Instance.LoadConfig(@"F:\ExcelGen\ExcelGen\bin\Debug\xml");
        }

    }


    public class PlayerActor : ScriptTableBase<PlayerActor, int>
    {
        // Fields
        private static PlayerActor _Instance = null;

        // Methods
        public static void TClear()
        {
            // Instance.Clear();
        }

        public static PlayerActor TGet(int Index)
        {
            return Instance.Get(Index);
        }

        // Properties
        [ExcelIndex]
        public int ID { get; set; }

        public new static PlayerActor Instance
        {
            get
            {
                if (_Instance == null)
                {
                    _Instance = new PlayerActor();
                }
                return _Instance;
            }
        }

        public NameObject Name { get; set; }

        public virtual string XMLFile
        {
            get
            {
                return "PlayerActor.xml";
            }
        }

        // Nested Types
        public class NameObject
        {
            // Properties
            public string Name { get; set; }

            public string NameDetail { get; set; }

            public int NameDetaili18n { get; set; }

            public int Namei18n { get; set; }
        }
    }

}
