using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Globalization;
using Microsoft.CSharp;
using System.CodeDom;
using System.CodeDom.Compiler;

namespace ExcelTool
{
    /// <summary>
    /// 动态编译
    /// </summary>
    class DynamicCompile
    {
        public Assembly Compile(string code, string outDll)
        {
            var options = new Dictionary<string, string> { { "CompilerVersion", "v4.0" } };
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider(options);

            // 2.ICodeComplier
            ICodeCompiler objICodeCompiler = objCSharpCodePrivoder.CreateCompiler();

            // 3.CompilerParameters
            CompilerParameters objCompilerParameters = new CompilerParameters();
            //objCompilerParameters.ReferencedAssemblies.Add("System.dll");
            objCompilerParameters.ReferencedAssemblies.Add(@"System.Xml.dll");
            objCompilerParameters.ReferencedAssemblies.Add(@"System.Xml.Linq.dll");
            // objCompilerParameters.ReferencedAssemblies.Add("Microsoft.CSharp.dll");
            objCompilerParameters.GenerateExecutable = false;
            objCompilerParameters.GenerateInMemory = false;
            objCompilerParameters.IncludeDebugInformation = true;

            objCompilerParameters.CompilerOptions = "/doc:" + outDll.Substring(0, outDll.LastIndexOf('.')) + ".xml";

            objCompilerParameters.OutputAssembly = outDll;

            // 4.CompilerResults
            CompilerResults cr = objICodeCompiler.CompileAssemblyFromSource(objCompilerParameters, code);

            if (cr.Errors.HasErrors)
            {
                Console.WriteLine("编译错误：");
                foreach (CompilerError err in cr.Errors)
                {
                    Console.WriteLine(err.ErrorText);
                }
                return null;
            }
            else
            {
                // 通过反射，调用HelloWorld的实例
                Assembly objAssembly = cr.CompiledAssembly;
                return objAssembly;
            }
        }
    }
}
