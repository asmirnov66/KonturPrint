using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using KonturPrint.Interfaces;
using Microsoft.CSharp;
using XMachine;

namespace ActiveMockLibrary
{
    public class CodeEvaluate
    {
        private const string CodeTemplate = @"
            using KonturPrint.Interfaces;
            using SKGENERALLib;

            namespace ActiveMockLibrary
            {
                public class Evaluator
                {
                    public void Eval(params object[] args)
                    {
                        #Code#
                    }
                }
            }
            ";
        public string Code { get; set; }
        public object[] Args { get; set; }

        public CodeEvaluate()
        {
            Code = "return;";
        }

        public CodeEvaluate(string code, object[] args)
        {
            Code = code;
            Args = args;
        }

        public bool Evaluate()
        {
            var source = CodeTemplate.Replace("#Code#", Code);
            var providerOptions = new Dictionary<string, string>
                {
                     {"CompilerVersion", "v3.5"}
                };
            var provider = new CSharpCodeProvider(providerOptions);
            var compilerParams = new CompilerParameters
            {
                GenerateInMemory = true,
                GenerateExecutable = false,
                CompilerOptions = "/platform:x86 /target:library"
            };
            var assemblies = Assembly.GetExecutingAssembly().DefinedTypes
                            .Select(t => t.Assembly.Location)
                            .Distinct()
                            .ToArray();
            compilerParams.ReferencedAssemblies.AddRange(assemblies);
            var references = new[] { "KonturPrint" };
            var konturPrintRef = AppDomain.CurrentDomain.GetAssemblies()
                                .Where(a => Array.Exists(references, el => string.Equals(el, a.GetName().Name, StringComparison.OrdinalIgnoreCase)))
                                .Select(a => a.Location)
                                .ToArray();
            compilerParams.ReferencedAssemblies.AddRange(konturPrintRef);
            var results = provider.CompileAssemblyFromSource(compilerParams, source);
            if (results.Errors.Count != 0)
            {
                return false;
            }
            var o = results.CompiledAssembly.CreateInstance("ActiveMockLibrary.Evaluator");
            if (o == null)
            {
                return false;
            }
            try
            {
                var mi = o.GetType().GetMethod("Eval");
                mi.Invoke(o, new object[] { Args });
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }

    public class XMachineMock : IXMachine
    {
        private CodeEvaluate CodeEvaluate { get; }

        public XMachineMock()
        {
            CodeEvaluate = new CodeEvaluate();
        }

        public object Evaluate(object source, params object[] args)
        {
            CodeEvaluate.Code = (string)source;
            CodeEvaluate.Args = args;
            return CodeEvaluate.Evaluate();
        }

        public object Call(string name, params object[] args)
        {
            throw new NotImplementedException();
        }

        public void LoadModule(string name)
        {
            throw new NotImplementedException();
        }

        public void UnloadModule(string name)
        {
            throw new NotImplementedException();
        }

        public void RegisterLib(string name)
        {
            throw new NotImplementedException();
        }

        public void RegisterVariable(string name, ref object _object)
        {
            throw new NotImplementedException();
        }

        public void ShowDebug()
        {
            throw new NotImplementedException();
        }

        public void EditScript()
        {
            throw new NotImplementedException();
        }

        public void EditCodeBlock()
        {
            throw new NotImplementedException();
        }

        public object CreateBlock(string source)
        {
            throw new NotImplementedException();
        }

        public void SetRoot(string name, object sourceRoot, object objectRoot, bool isDefault)
        {
            throw new NotImplementedException();
        }

        public object GetGlobalNames()
        {
            throw new NotImplementedException();
        }

        public void RegisterImport(string name)
        {
            throw new NotImplementedException();
        }

        public int TestMethod(int x)
        {
            throw new NotImplementedException();
        }

        public int TestMethodEnum(TestEnum x)
        {
            throw new NotImplementedException();
        }

        public int TestMethodArray(ref Array array)
        {
            throw new NotImplementedException();
        }

        public int TestMethodPtr(ref int x)
        {
            throw new NotImplementedException();
        }
    }
}
