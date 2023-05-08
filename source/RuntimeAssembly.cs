using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.CSharp;

namespace ExcelToJson
{
    public static class RuntimeAssembly
    {
        public static Assembly RuntimeAsm;

        private static string BaseProgram = @"
using System.Collections.Generic;
using System;
            class Program{
                static void Main(string[] args){
                }
            }";
        
        public static void Compile(List<string> classes)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string @class in classes)
            {
                sb.Append(@class);
            }

            var success = CompileAssembly(sb.ToString(), out RuntimeAsm, out var error);
            if (success) return;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("compile error！！！");
            Console.WriteLine(error);
            Console.ResetColor();
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
            references.Add(MetadataReference.CreateFromFile(typeof(MongoDB.Bson.Serialization.BsonSerializer).Assembly.Location));

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
    }
}