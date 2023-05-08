using System;
using System.IO;
using System.Text;
using CommandLine;

namespace ExcelToJson
{
    /// <summary>
    /// 命令行参数定义
    /// </summary>
    public sealed class Options
    {
        public static Options Default = new Options();

        public Encoding Encoding => new UTF8Encoding(false);

        [Option('e', "excel_path")] public string ExcelPath { get; set; }
        [Option('j', "json_path")] public string JsonPath { get; set; }
        [Option('c', "script_path")] public string ScriptPath { get; set; }
        [Option('t', "script_template_path")] public string ScriptTemplate { get; set; }

        [Option('l', "script_load_json")] 
        public string ScriptLoadJsonPath { get; set; } = "#json_name";

        [Option('b',"bson")]
        public int IsBinaryJson { get; set; }

        public string JsonFileExtension => IsBinaryJson == 1 ? ".bytes" : ".json";

        public void Check()
        {
            ExcelPath = Path.GetFullPath(ExcelPath);
            ScriptTemplate = Path.GetFullPath(ScriptTemplate);
            if (Path.HasExtension(ScriptLoadJsonPath))
                ScriptLoadJsonPath = ScriptLoadJsonPath.Replace(Path.GetExtension(ScriptLoadJsonPath), "");

            if (!Directory.Exists(ExcelPath))
                Console.WriteLine($"{ExcelPath} not exist");

            if (string.IsNullOrEmpty(ScriptPath))
            {
                Console.WriteLine("ScriptPath is null");
            }
            else
            {
                ScriptPath = Path.GetFullPath(ScriptPath);
            }

            if (string.IsNullOrEmpty(JsonPath))
            {
                Console.WriteLine("JsonPath is null");
            }
            else
            {
                JsonPath = Path.GetFullPath(JsonPath);
            }
        }
    }
}