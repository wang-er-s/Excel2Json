using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToJson
{

    sealed class Program
    {
        private static List<string> excelPaths = new List<string>();
        private static List<string> excelFileNames = new List<string>();
        private static readonly object lockObj = new object();
        private static bool finish = false;

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine(string.Join(" ", args));
            Console.ForegroundColor = ConsoleColor.Cyan;
            Stopwatch sw = Stopwatch.StartNew();
            sw.Start();
            ParseConfig(args);
            CreateDir();

            //第一次CreateOpenXmlReader需要在单一线程跑，否则会出现错误
            foreach (var excelPath in excelPaths)
            {
                Tools.CreateExcelReader(excelPath);
            }
            Gen();
            while (!finish)
            {
                Thread.Sleep(100);
            }
            sw.Stop();
            Console.WriteLine($"use {sw.ElapsedMilliseconds}ms, press any key to exit");
            Console.ResetColor();
            Console.ReadKey();
        }
        
        private static async void Gen()
        {
            GenEnum();
            List<string> cs = new List<string>();
            int length = excelPaths.Count;
            List<Task> allTask = new List<Task>(excelPaths.Count);
            int index = 0;
            foreach (string excelPath in excelPaths)
            {
                allTask.Add(Task.Run(() =>
                {
                    var str = CSExporter.GenClass(excelPath);
                    RuntimeAssembly.CheckClass(str);
                    lock (lockObj)
                    {
                        cs.Add(str);
                        index++;
                        Console.WriteLine($"gen {Path.GetFileNameWithoutExtension(excelPath)} cs...({index}/{length})");
                    }
                }));
            }
            await Task.WhenAll(allTask);

            RuntimeAssembly.Compile(cs);
            index = 0;
            allTask.Clear();
            foreach (string excelPath in excelPaths)
            {
                allTask.Add(Task.Run(() =>
                {
                    JsonExporter.GenJson(excelPath);
                    lock (lockObj)
                    {
                        index++;
                        Console.WriteLine($"gen {Path.GetFileNameWithoutExtension(excelPath)} json...({index}/{length})");
                    }
                }));
            }
            await Task.WhenAll(allTask);
            finish = true;
            Console.WriteLine(" ************** Export Succeed ! *************** ");
        }

        private static void GenSync()
        {
            GenEnum();
            List<string> cs = new List<string>();
            int length = excelPaths.Count;
            int index = 0;
            foreach (string excelPath in excelPaths)
            {

                var str = CSExporter.GenClass(excelPath);
                cs.Add(str);
                index++;
                Console.WriteLine($"gen {Path.GetFileNameWithoutExtension(excelPath)} cs...({index}/{length})");
            }

            RuntimeAssembly.Compile(cs);
            index = 0;
            foreach (string excelPath in excelPaths)
            {
                JsonExporter.GenJson(excelPath);
                index++;
                Console.WriteLine($"gen {Path.GetFileNameWithoutExtension(excelPath)} json...({index}/{length})");
            }
            finish = true;
            Console.WriteLine(" ************** Export Succeed ! *************** ");
        }

        private static void GenEnum()
        {
            StringBuilder enumSb = new StringBuilder();
            foreach (string excelPath in excelPaths)
            {
                var enumStr = CSExporter.GenEnum(excelPath);
                enumSb.AppendLine(enumStr);
            }

            RuntimeAssembly.AddToBaseProgram(enumSb.ToString());
        }

        private static void ParseConfig(string[] args)
        {
            Options.Default = CommandLine.Parser.Default.ParseArguments<Options>(args).Value;
            Options.Default.Check();
            DirectoryInfo root = new DirectoryInfo(Options.Default.ExcelPath);
            var files = root.GetFiles("*.xls").Concat(root.GetFiles("*.xlsx"));
            files = files.Distinct();
            files = files.Where((info => !info.Name.StartsWith("~")));
            foreach (FileInfo file in files)
            {
                if (excelPaths.Contains(file.FullName)) continue;
                excelPaths.Add(file.FullName);
                var fileName = Path.GetFileNameWithoutExtension(file.Name);
                if (!Regex.Match(fileName, @"^[a-zA-Z][\w_]+").Success)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"{fileName} 表格名字不符合规范，需要以字母开头，且名字中只能包含字母数字下划线");
                    throw new Exception();
                }
            }
        }

        private static void CreateDir()
        {
            if (!string.IsNullOrEmpty(Options.Default.ScriptPath))
            {
                if (Directory.Exists(Options.Default.ScriptPath))
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(Options.Default.ScriptPath);
                    foreach (var fileInfo in directoryInfo.GetFiles())
                    {
                        fileInfo.Delete();
                    }
                }
                else
                {
                    Directory.CreateDirectory(Options.Default.ScriptPath);
                }
            }
            if (!string.IsNullOrEmpty(Options.Default.JsonPath))
            {
                if (Directory.Exists(Options.Default.JsonPath))
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(Options.Default.JsonPath);
                    foreach (var fileInfo in directoryInfo.GetFiles())
                    {
                        fileInfo.Delete();
                    }
                }
                else
                {
                    Directory.CreateDirectory(Options.Default.JsonPath);
                }
            }
        }
    }
}

