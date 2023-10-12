using System;
using System.IO;
using System.Text;
using CommandLine;

namespace ExcelToJson
{
    /// <summary>
    ///     命令行参数定义
    /// </summary>
    public sealed class Options
    {
        public static Options Default = new();

        public Encoding Encoding => new UTF8Encoding(false);

        [Option("excel_path")]
        public string ExcelPath { get; set; }

        [Option("json_path")]
        public string JsonPath { get; set; }

        [Option("script_path")]
        public string ScriptPath { get; set; }

        [Option("script_template_path")]
        public string ScriptTemplate { get; set; }

        public void Check()
        {
            if (!Directory.Exists(ExcelPath))
            {
                Console.WriteLine($"{ExcelPath} not exist");
            }

            if (string.IsNullOrEmpty(ScriptPath))
            {
                Console.WriteLine("ScriptPath is null");
            }
            else
            {
                if (!File.Exists(ScriptTemplate))
                {
                    Console.WriteLine($"ScriptTemplatePath:{ScriptTemplate} not exist");
                }
            }

            if (string.IsNullOrEmpty(JsonPath))
            {
                Console.WriteLine("JsonPath is null");
            }
        }
    }
}