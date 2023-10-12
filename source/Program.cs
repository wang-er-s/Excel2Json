using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelToJson
{
    internal sealed class Program
    {
        private static readonly object lockObj = new();
        public static List<string> ExportExcelPaths = new();

        [STAThread]
        private static void Main(string[] args)
        {
            ParseConfig(args);
            Stopwatch sw = Stopwatch.StartNew();
            sw.Start();
            GenSync();
            sw.Stop();
            Console.WriteLine($"use {sw.ElapsedMilliseconds}ms");
        }

        private static void ParseConfig(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments(args, Options.Default);
            Options.Default.Check();
            var files = Tools.GetExcelFiles(Options.Default.ExcelPath);
            foreach (FileInfo file in files)
            {
                if (ExportExcelPaths.Contains(file.FullName)) continue;
                ExportExcelPaths.Add(file.FullName);
            }
        }

        private static void GenSync()
        {
            PreCompileEnum();
            Dictionary<string, OneClassData> data = new();
            List<string> cs = new();
            foreach (string excelPath in ExportExcelPaths)
            {
                string str = ClassGenerator.GenClass(excelPath, data, out bool check);
                if (check)
                {
                    RuntimeAssembly.CheckClass(str);
                }

                Console.WriteLine($"{Path.GetFileName(excelPath)}....");

                lock (lockObj)
                {
                    cs.Add(str);
                }
            }

            RuntimeAssembly.Compile(cs);

            foreach (string excelPath in ExportExcelPaths)
            {
                DataFiller.FillData(excelPath, data);
            }

            switch (Options.Default.ScriptType)
            {
                case ScriptType.CS:
                    CSExporter.Export(data);
                    break;
                case ScriptType.TS:
                    TsExporter.Export(data);
                    break;
            }
        }

        private static void PreCompileEnum()
        {
            RuntimeAssembly.Clear();

            StringBuilder enumSb = new();

            int index = 0;


            Directory.CreateDirectory(Options.Default.ScriptPath);
            Directory.CreateDirectory(Options.Default.JsonPath);
            var excelFile  = Tools.GetExcelFiles(Options.Default.ExcelPath);
            foreach (var file in excelFile)
            {
                try
                {
                    string enumStr = ClassGenerator.GenEnum(file.FullName, index);
                    index++;
                    enumSb.AppendLine(enumStr);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            RuntimeAssembly.AddToBaseProgram(enumSb.ToString());
        }
    }
}