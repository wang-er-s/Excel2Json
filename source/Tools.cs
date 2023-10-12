using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelToJson
{
    public static class Tools
    {
        public const string EnumTableName = "enum";
        public const int HeaderRows = 3;
        public const string ExcelEndFlag = "//END";
        //忽略行或列的前缀
        public const string IgnorePrefix = "#";

        /// <summary>
        /// 根据表名获取生成的类或者json的文件名
        /// </summary>
        public static string GetTableGenName(string excelPath, string tableName)
        {
            string name = tableName;
            if (Regex.Match(tableName, @"Sheet\d+", RegexOptions.IgnoreCase).Success)
            {
                name = Path.GetFileNameWithoutExtension(excelPath);
            }
            else
            {
                name = tableName;
                var className = Path.GetFileNameWithoutExtension(excelPath);
                if (Program.ExportExcelPaths.Contains(className))
                    Program.ExportExcelPaths.Add(name);
            }
            return name;
        }

        public static bool CheckTableIsEnum(string tableName)
        {
            return tableName.ToLower() == EnumTableName;
        }
        
        public static bool CheckTableIsIgnore(string tableName)
        {
            return (tableName.StartsWith("#") || tableName == "说明");
        }

        public static List<FileInfo> GetExcelFiles(string path)
        {
            List<FileInfo> result = new List<FileInfo>();
            DirectoryInfo root = new DirectoryInfo(path);
            var files = root.GetFiles("*.xls");
            foreach (FileInfo fileInfo in files)
            {
                if (!fileInfo.Name.StartsWith("~") && Path.GetExtension(fileInfo.Name) == ".xls")
                {
                    result.Add(fileInfo);
                }
            }
            
            files = root.GetFiles("*.xlsx");
            foreach (FileInfo fileInfo in files)
            {
                if (!fileInfo.Name.StartsWith("~") &&  Path.GetExtension(fileInfo.Name) == ".xlsx")
                {
                    result.Add(fileInfo);
                }
            }

            return result;
        }
    }
}