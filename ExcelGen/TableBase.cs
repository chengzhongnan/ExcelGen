using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml.Linq;

namespace ExcelTool
{

    public class IScriptTableBase
    {
    }

    public class ExcelIndexAttribute : Attribute
    {

    }

    /// <summary>
    /// 配置文件类基类
    /// </summary>
    /// <typeparam name="T">子类类型</typeparam>
    /// <typeparam name="U">索引类型</typeparam>
    public class ScriptTableBase<T,U> : IScriptTableBase
        where T : ScriptTableBase<T,U>, new()
    {
        private static T _Instance = null;
        /// <summary>
        /// 静态调用实例
        /// </summary>
        public static T Instance
        {
            get
            {
                if(_Instance == null)
                {
                    _Instance = new T();
                }
                return _Instance;
            }
        }
        private Dictionary<U, T> _Tables = null;
        /// <summary>
        /// 静态配置字典
        /// </summary>
        public Dictionary<U, T> Tables
        {
            get
            {
                if(_Tables == null)
                {
                    _Tables = new Dictionary<U, T>();
                }
                return _Tables;
            }
        }

        /// <summary>
        /// 添加数据到Table字典中
        /// </summary>
        /// <param name="table"></param>
        public virtual void Add(T table)
        {
            var props = typeof(T).GetProperties();
            PropertyInfo indexProp = null;
            foreach(var p in props)
            {
                if(p.PropertyType != typeof(U))
                {
                    continue;
                }
                foreach(var pAttr in p.CustomAttributes)
                {
                    if(pAttr.AttributeType == typeof(ExcelIndexAttribute))
                    {
                        indexProp = p;
                        break;
                    }
                }
                if(indexProp != null)
                {
                    break;
                }
            }
            if(indexProp != null)
            {
                var indexValue = indexProp.GetValue(table);
                Tables.Add((U)indexValue, table);
            }
        }

        /// <summary>
        /// 从Table中获取数据
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public virtual T Get(U Index)
        {
            if (Tables.ContainsKey(Index))
            {
                return Tables[Index];
            }
            return default(T);
        }

        /// <summary>
        /// 加载配置文件
        /// </summary>
        /// <param name="fileName"></param>
        public virtual void LoadTable(string fileName)
        {
            Tables.Clear();
            var xRoot = XElement.Load(fileName);
            foreach(var xTable in xRoot.Elements())
            {
                try
                {
                    T table = LoadTableFromXmlNode(xTable);
                    Add(table);
                }
                catch (Exception ex)
                {
                    Exception propException = new Exception(
                        string.Format("TableType=[{0}], fileName = {1}", typeof(T), fileName), ex);
                    throw propException;
                }
            }
        }

        protected T LoadTableFromXmlNode(XElement xEle)
        {
            var type = typeof(T);
            var propList = type.GetProperties();
            T objTable = new T();
            foreach(var pro in propList)
            {
                try
                {
                    if (pro.PropertyType.IsGenericType)
                    {
                        var propName = pro.Name;
                        if (propName.EndsWith("List"))
                        {
                            propName = propName.Substring(0, propName.Length - 4);
                        }
                        var xProp = xEle.Element(propName);
                        if (xProp != null)
                        {
                            var v = GetXmlListObject(xProp, pro);
                            pro.SetValue(objTable, v);
                        }
                    }
                    else
                    {
                        var xProp = xEle.Element(pro.Name);
                        if (xProp != null)
                        {
                            var sValue = xProp.Value;
                            object v = Convert.ChangeType(sValue, pro.PropertyType);
                            pro.SetValue(objTable, v);
                        }
                    }
                }
                catch(Exception ex)
                {
                    Exception propException = new Exception(
                        string.Format("TableType=[{0}], PropName = [{1}], xml = [{2}]", typeof(T), pro.Name, xEle.ToString()), ex);
                    throw propException;
                }
            }
            return objTable;
        }

        private object GetXmlListObject(XElement xProp, PropertyInfo propInfo)
        {
            var genericArgv = propInfo.PropertyType.GenericTypeArguments[0];
            dynamic obj = Activator.CreateInstance(propInfo.PropertyType);
            var addMethod = propInfo.PropertyType.GetMethod("Add");
            if (genericArgv.IsPrimitive)
            {
                // 基础类型
                foreach (var subNode in xProp.Elements())
                {
                    var s = subNode.Value;
                    object v = Convert.ChangeType(s, genericArgv);
                    addMethod.Invoke(obj, new object[] { v });
                }
            }
            else
            {
                // 非基础类型
                foreach(var subNode in xProp.Elements())
                {
                    var objSub = Activator.CreateInstance(genericArgv);
                    var objSubProps = objSub.GetType().GetProperties();
                    foreach(var subObjProp in objSubProps)
                    {
                        var s = subNode.Element(subObjProp.Name);
                        if(s != null)
                        {
                            object v = Convert.ChangeType(s.Value, subObjProp.PropertyType);
                            subObjProp.SetValue(objSub, v);
                        }
                    }

                    addMethod.Invoke(obj, new object[] { objSub });
                }
            }
            return obj;
        }
    }

    public class AutoLoadConfig
    {
        public AutoLoadConfig()
        {
            TableConfigTypeList = new Dictionary<string, Type>();
            RegistryAssembly(typeof(AutoLoadConfig).Assembly);
        }

        private static AutoLoadConfig _Instance = null;
        public static AutoLoadConfig Instance
        {
            get
            {
                if(_Instance == null)
                {
                    _Instance = new AutoLoadConfig();
                }
                return _Instance;
            }
        }

        private Dictionary<string ,Type> TableConfigTypeList = null; 

        public void RegistryAssembly(Assembly asm)
        {
            var types = asm.GetTypes();
            foreach(var type in types)
            {
                if(type.IsSubclassOf(typeof(IScriptTableBase)) && !type.Name.StartsWith("ScriptTableBase"))
                {
                    if(!TableConfigTypeList.ContainsKey(type.AssemblyQualifiedName))
                    {
                        TableConfigTypeList.Add(type.AssemblyQualifiedName, type);
                    }
                }
            }
        }

        public void LoadConfig(string strFolder)
        {
            foreach(var typeKV in TableConfigTypeList)
            {
                try
                {
                    var instanceProp = typeKV.Value.GetProperty("Instance", BindingFlags.Static | BindingFlags.Public);
                    if (instanceProp != null)
                    {
                        var instance = instanceProp.GetValue(null);

                        var xmlProp = typeKV.Value.GetProperty("XMLFile", BindingFlags.Instance | BindingFlags.Public);
                        string xmlName = "";
                        if (xmlProp != null)
                        {
                            xmlName = xmlProp.GetValue(instance).ToString();
                        }
                        else
                        {
                            xmlName = typeKV.Value.Name + ".xml";
                        }

                        var xmlFileName = strFolder.TrimEnd('\\') + "\\" + xmlName;
                        var loadTableMethod = typeKV.Value.GetMethod("LoadTable");
                        if (instance != null && loadTableMethod != null)
                        {
                            loadTableMethod.Invoke(instance, new object[] { xmlFileName });
                        }
                    }
                }
                catch(Exception ex)
                {

                }
            }
        }
    }
}
