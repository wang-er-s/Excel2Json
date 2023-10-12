using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using Microsoft.CSharp;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis;
using System.IO;
using Microsoft.CodeAnalysis.Emit;
using System.Linq;

namespace ExcelToJson
{
    public static class RuntimeAssembly
    {
        public static Assembly RuntimeAsm;
        private readonly static string Template = @"
using System.Collections.Generic;
using System;
            class Program{
                static void Main(string[] args){
                }
            }";

        private static string BaseProgram = Template;

        public static void Compile(List<string> classes)
        {
            var csc = new CSharpCodeProvider(new Dictionary<string, string>() {{"CompilerVersion", "v3.5"}});
            var parameters = new CompilerParameters(new[] {"mscorlib.dll", "System.Core.dll"}, "", false);
            parameters.GenerateExecutable = true;
            StringBuilder sb = new StringBuilder();
            sb.Append(BaseProgram);
            foreach (string @class in classes)
            {
                sb.Append(@class);
            }
            CompilerResults results = csc.CompileAssemblyFromSource(parameters, sb.ToString());
            sb.Clear();
            if (results.Errors.HasErrors)
            {
                foreach (CompilerError compilerError in results.Errors)
                {
                    sb.AppendLine(compilerError.ErrorText);
                }
                Console.WriteLine("compile error！！！");
                Console.WriteLine(sb.ToString());
            }
            RuntimeAsm = results.CompiledAssembly;
        }
        
        public static void AddToBaseProgram(string str)
        {
            BaseProgram += str;
        }
        
        public static void CheckClass(string @class)
        {
            var success = CompileAssembly(@class, out var _, out string error);
            if (success) return;
            var match = Regex.Match(@class, @"class\s+([\w\W]*?)\{");
            string typeName = String.Empty;
            if (match.Success)
            {
                typeName = match.Groups[1].Value;
            }
            else
            {
                Console.WriteLine("???为什么会匹配失败");
            }

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($">>>{typeName}<<< compile error！！！");
            Console.WriteLine(error);
            Console.ResetColor();
            throw new Exception(error);
        }
        
        private static bool CompileAssembly(string classContent, out Assembly resultAssembly, out string error)
        {
            resultAssembly = null;
            error = string.Empty;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(BaseProgram);
            sb.AppendLine(classContent);
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(sb.ToString());

            // define other necessary objects for compilation
            string assemblyName = Path.GetRandomFileName();
            List<MetadataReference> references = new List<MetadataReference>();
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (Assembly assembly in assemblies)
            {
                try
                {
                    if (assembly.IsDynamic)
                    {
                        continue;
                    }

                    if (assembly.Location == "")
                    {
                        continue;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

                PortableExecutableReference reference = MetadataReference.CreateFromFile(assembly.Location);
                references.Add(reference);
            }

            // analyse and generate IL code from syntax tree
            CSharpCompilation compilation = CSharpCompilation.Create(
                assemblyName,
                syntaxTrees: new[] { syntaxTree },
                references: references,
                options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            using var ms = new MemoryStream();
            // write IL code into memory
            EmitResult result = compilation.Emit(ms);

            if (result.Success)
            {
                resultAssembly = Assembly.Load(ms.ToArray());
                return true;
            }

            // handle exceptions
            IEnumerable<Diagnostic> failures = result.Diagnostics.Where(diagnostic => 
                diagnostic.IsWarningAsError || 
                diagnostic.Severity == DiagnosticSeverity.Error);

            sb.Clear();
            foreach (Diagnostic diagnostic in failures)
            {
                sb.AppendLine(diagnostic.GetMessage());
            }

            error = sb.ToString();
            return false;
        }

        public static void Clear()
        {
            BaseProgram = Template;
        }
    }
}