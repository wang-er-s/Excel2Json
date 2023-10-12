using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Excel;

namespace ExcelToJson
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    static class ClassGenerator
    {
        struct FieldDef
        {
            public string Name;
            public string Type;
            public string Remark;
        }

        private const string arrayPre = "array_";
        private const string dicPre = "dic_";
        private const string array2D = "array_two_";


        public static string GenClass(string excelPath, Dictionary<string, OneClassData> classDatas, out bool needCheck)
        {
            needCheck = false;
            // 加载Excel文件
            var extension = Path.GetExtension(excelPath);
            if (extension != ".xls" && extension != ".xlsx")
            {
                throw new Exception("unSupport excel type " + extension);
            }

            IExcelDataReader excelReader;
            using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
            }

            // 数据检测
            if (excelReader.ResultsCount < 1)
            {
                throw new Exception("Excel file is empty: " + excelPath);
            }

            StringBuilder result = new StringBuilder();
            try
            {
                do
                {
                    var excelName = excelReader.Name;
                    if (Tools.CheckTableIsEnum(excelName))
                    {
                        continue;
                    }

                    if (Tools.CheckTableIsIgnore(excelName)) continue;
                    OneClassData classData = new OneClassData();
                    result.Append(GenOneTableCompileCode(excelReader, excelPath, classData));
                    classDatas[classData.ClassName] = classData;
                    needCheck = true;
                } while (excelReader.NextResult());
            }
            catch (Exception e)
            {
                Console.WriteLine($"table {Path.GetFileNameWithoutExtension(excelPath)} parse cs error !!!! \n{e}");
                throw;
            }

            return result.ToString();
        }

        private static string GenOneTableCompileCode(IExcelDataReader excelReader, string excelPath,
            OneClassData oneClassData)
        {
            StringBuilder compileCs = new StringBuilder("public class #class_name{#fields}");
            string csName = Tools.GetTableGenName(excelPath, excelReader.Name);
            var fieldList = ParseField(excelReader);
            if (fieldList.Count <= 0) return "";
            compileCs = compileCs.Replace("#class_name", csName);
            oneClassData.ClassName = csName;
            oneClassData.FileName = Path.GetFileNameWithoutExtension(excelPath);
            //Console.WriteLine(oneClassData.FileName);
            StringBuilder sb = new StringBuilder();
            foreach (FieldDef field in fieldList)
            {
                sb.AppendLine($"/// <summary> {field.Remark} </summary>");
                sb.AppendFormat("\tpublic {0} {1} {{ get; private set; }}", field.Type, field.Name);
                sb.AppendLine();
                oneClassData.PropName2Comment[field.Name] = field.Remark;
            }

            compileCs = compileCs.Replace("#fields", sb.ToString());
            return compileCs.ToString();
        }

        private static List<FieldDef> ParseField(IExcelDataReader excelReader)
        {
            var result = new List<FieldDef>();
            int row = 0;
            List<string> names = new List<string>();
            List<string> types = new List<string>();
            List<string> comments = new List<string>();
            List<int> ignoreLine = new List<int>();
            // 第一行是字段名 第二行是类型 第三行是注释
            while (excelReader.Read())
            {
                List<string> targetList;
                row++;
                if (row == 1)
                {
                    targetList = names;
                }
                else if (row == 2)
                {
                    targetList = types;
                }
                else if (row == 3)
                {
                    targetList = comments;
                }
                else
                {
                    break;
                }

                for (int i = 0; i < excelReader.FieldCount; i++)
                {
                    if (ignoreLine.Contains(i)) continue;
                    var str = excelReader.GetString(i) ?? string.Empty;
                    //如果是第一行，而且是#开头 则忽略
                    if (row == 1 && str.StartsWith(Tools.IgnorePrefix))
                    {
                        ignoreLine.Add(i);
                        continue;
                    }

                    targetList.Add(str);
                }
            }

            for (int i = 0; i < names.Count; i++)
            {
                var name = names[i];
                var type = types[i];
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(type)) break;
                var comment = comments[i];
                FieldDef field;
                field.Name = name.Trim();
                if (type.Contains(array2D))
                {
                    var listTypeStr = type.Replace(array2D, "");
                    type = $"{listTypeStr}[][]";
                }
                else if (type.Contains(arrayPre))
                {
                    var listTypeStr = type.Replace(arrayPre, "");
                    type = $"List<{listTypeStr}>";
                }
                else if (type.Contains(dicPre))
                {
                    var match = Regex.Match(type, @"\w+_(\w+)_(\w+)");
                    type = $"Dictionary<{match.Groups[1].Value}, {match.Groups[2].Value}>";
                }
                else if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(field.Name))
                {
                    continue;
                }

                field.Type = type;
                field.Remark = Regex.Replace(comment, "[\r\n\t]", " ", RegexOptions.Compiled);
                result.Add(field);
            }

            return result;
        }

        public static string GenEnum(string excelPath, int index)
        {
            // 加载Excel文件
            IExcelDataReader excelReader;
            using (FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);
            }

            // 数据检测
            if (excelReader.ResultsCount < 1)
            {
                throw new Exception("Excel file is empty: " + excelPath);
            }

            StringBuilder result = new StringBuilder();
            try
            {
                do
                {
                    var excelName = excelReader.Name;
                    if (Tools.CheckTableIsEnum(excelName))
                    {
                        result.Append(GenEnum(excelReader, excelPath, index));
                    }
                } while (excelReader.NextResult());
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"table [[{Path.GetFileNameWithoutExtension(excelPath)}]] parse cs error !!!! \n{e}");
                Console.ResetColor();
                throw;
            }

            return result.ToString();
        }

        class KeyValue<TKey, TValue>
        {
            public TKey Key;
            public TValue Value;
        }

        private static string GenEnum(IExcelDataReader excelReader, string excelPath, int index)
        {
            int row = 0;
            List<StringBuilder> stringBuilders = new List<StringBuilder>();
            Dictionary<string, List<KeyValue<string, int>>>
                data = new Dictionary<string, List<KeyValue<string, int>>>();
            string enumTypeName = String.Empty;
            while (excelReader.Read())
            {
                if (stringBuilders.Count <= 0)
                {
                    stringBuilders = new List<StringBuilder>();
                    for (int i = 0; i < excelReader.FieldCount; i++)
                    {
                        var obj = excelReader.GetString(i);
                        if (string.IsNullOrEmpty(obj)) continue;
                        stringBuilders.Add(new StringBuilder());
                    }
                }

                // i是列 row是行
                row++;
                if (row <= 3) continue;
                for (int i = 0; i < 3; i++)
                {
                    var obj = excelReader.GetString(i);
                    if (i == 0 && !string.IsNullOrEmpty(obj))
                    {
                        data[obj] = new List<KeyValue<string, int>>();
                        stringBuilders[i].AppendLine($"public enum {obj}{{");
                        enumTypeName = obj;
                    }
                    else if (i == 1 && !string.IsNullOrEmpty(obj))
                    {
                        data[enumTypeName].Add(new KeyValue<string, int>() { Key = obj, Value = -1 });
                    }
                    else if (i == 2 && !string.IsNullOrEmpty(obj))
                    {
                        try
                        {
                            int val = Int32.Parse(obj.Trim());

                            foreach (var item in data[enumTypeName])
                            {
                                if (item.Value == val)
                                {
                                    throw new Exception($"{excelPath} enum {enumTypeName} has same value [{val}]");
                                }
                            }

                            data[enumTypeName][data[enumTypeName].Count - 1].Value = val;
                        }
                        catch (Exception _)
                        {
                            Console.WriteLine($"{excelPath} enum {enumTypeName} has wrong value");
                            throw;
                        }
                    }
                }
            }

            StringBuilder result = new StringBuilder();
            foreach (var item in data)
            {
                result.AppendLine($"public enum {item.Key}{{");
                foreach (var keyvalue in item.Value)
                {
                    if (keyvalue.Value != -1)
                    {
                        result.AppendLine($"\t{keyvalue.Key} = {keyvalue.Value},");
                    }
                    else
                    {
                        result.AppendLine($"\t{keyvalue.Key},");
                    }
                }

                result.AppendLine("}");
            }

            return result.ToString();
        }
    }
}