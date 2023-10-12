using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;

namespace ExcelToJson
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    static class CSExporter
    {
        public static void Export(Dictionary<string, OneClassData> data)
        {
            if (Directory.Exists(Options.Default.ScriptPath))
            {
                string[] files = Directory.GetFiles(Options.Default.ScriptPath);
                foreach (string file in files)
                {
                    if (file.EndsWith(".cs"))
                        File.Delete(file);
                }
            } 
            string scriptTemplate = File.ReadAllText(Options.Default.ScriptTemplate);
            StringBuilder enumSb = new StringBuilder();
            foreach (var type in RuntimeAssembly.RuntimeAsm.GetTypes())
            {
                if (type.Name == "Program")
                    continue;
                if (type.IsEnum)
                {
                    enumSb.AppendLine($"public enum {type.Name}{{");
                    foreach (string enumName in Enum.GetNames(type))
                    {
                        int enumValue = Convert.ToInt32(Enum.Parse(type, enumName));
                        if (enumValue == -1)
                        {
                            enumSb.AppendLine($"\t{enumName},");
                        }
                        else
                        {
                            enumSb.AppendLine($"\t{enumName}= {enumValue},");
                        }
                    }

                    enumSb.AppendLine("}");
                }
                else
                {
                    StringBuilder sb = new StringBuilder(scriptTemplate);
                    StringBuilder filedSb = new StringBuilder();
                    // 生成ts文件
                    if (!string.IsNullOrEmpty(Options.Default.ScriptPath))
                    {
                        var properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                        foreach (PropertyInfo property in properties)
                        {
                            if (data[type.Name].PropName2Comment.TryGetValue(property.Name, out string value))
                            {
                                filedSb.AppendLine($"\t/// {value}");
                            }

                            if (typeof(IDictionary).IsAssignableFrom(property.PropertyType))
                            {
                                var keyGenericType = property.PropertyType.GetGenericArguments()[0];
                                var valueGenericType = property.PropertyType.GetGenericArguments()[1];
                                filedSb.AppendLine($"\tpublic Dictionary<{keyGenericType},{valueGenericType}> {property.Name}{{get; private set;}}");
                            }
                            else if (typeof(IList).IsAssignableFrom(property.PropertyType))
                            {
                                var genericType = property.PropertyType.GetGenericArguments()[0];
                                filedSb.AppendLine($"\tpublic List<{genericType}> {property.Name}{{get; private set;}}");
                            }
                            else
                            {
                                filedSb.AppendLine($"\tpublic {property.PropertyType.Name} {property.Name}{{get; private set;}}");
                            }
                        }

                        sb.Replace("#class_name", type.Name);
                        sb.Replace("#fields", filedSb.ToString());
                        File.WriteAllText(Path.Combine(Options.Default.ScriptPath, type.Name + ".cs"),
                            sb.ToString());
                    }

                    // 生成json文件
                    File.WriteAllText(Path.Combine(Options.Default.JsonPath, type.Name + ".json"),
                        JsonConvert.SerializeObject(data[type.Name].Data, Formatting.Indented));
                }
            }

            if (enumSb.Length > 0)
            {
                File.WriteAllText(Path.Combine(Options.Default.ScriptPath, "enum.cs"), enumSb.ToString());
            }
        }
    }
}