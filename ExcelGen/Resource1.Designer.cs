﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:4.0.30319.42000
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace ExcelTool {
    using System;
    
    
    /// <summary>
    ///   一个强类型的资源类，用于查找本地化的字符串等。
    /// </summary>
    // 此类是由 StronglyTypedResourceBuilder
    // 类通过类似于 ResGen 或 Visual Studio 的工具自动生成的。
    // 若要添加或移除成员，请编辑 .ResX 文件，然后重新运行 ResGen
    // (以 /str 作为命令选项)，或重新生成 VS 项目。
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resource1 {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resource1() {
        }
        
        /// <summary>
        ///   返回此类使用的缓存的 ResourceManager 实例。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("ExcelTool.Resource1", typeof(Resource1).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   重写当前线程的 CurrentUICulture 属性
        ///   重写当前线程的 CurrentUICulture 属性。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   查找类似 using System;
        ///using System.Collections;
        ///using System.Collections.Generic;
        ///using System.Text;
        ///using System.Reflection;
        ///using System.Xml.Linq;
        ///
        ///namespace ExcelTables
        /// 的本地化字符串。
        /// </summary>
        internal static string ClassTemplete {
            get {
                return ResourceManager.GetString("ClassTemplete", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 public class IScriptTableBase
        ///    {
        ///    }
        ///
        ///    public class ExcelIndexAttribute : Attribute
        ///    {
        ///
        ///    }
        ///
        ///    /// &lt;summary&gt;
        ///    /// 配置文件类基类
        ///    /// &lt;/summary&gt;
        ///    /// &lt;typeparam name=&quot;T&quot;&gt;子类类型&lt;/typeparam&gt;
        ///    /// &lt;typeparam name=&quot;U&quot;&gt;索引类型&lt;/typeparam&gt;
        ///    public class ScriptTableBase&lt;T,U&gt; : IScriptTableBase
        ///        where T : ScriptTableBase&lt;T,U&gt;, new()
        ///    {
        ///        private static T _Instance = null;
        ///        /// &lt;summary&gt;
        ///        /// 静态调用实例
        ///        /// &lt;/summary&gt;
        ///        public static T In [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        internal static string ScriptBase {
            get {
                return ResourceManager.GetString("ScriptBase", resourceCulture);
            }
        }
    }
}
