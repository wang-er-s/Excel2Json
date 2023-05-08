using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using Excel;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Serializers;

namespace ExcelToJson
{
    static class JsonExporter
    {
        public static void GenJson(string excelPath)
        {
            try
            {
                IExcelDataReader excelReader = Tools.CreateExcelReader(excelPath);
                do
                {
                    var excelName = excelReader.Name;
                    if (Tools.CheckTableIsEnum(excelName)) continue;
                    if (Tools.CheckTableIsIgnore(excelName)) continue;
                    if (!string.IsNullOrEmpty(Options.Default.JsonPath))
                    {
                        var typeName = Tools.GetTableGenName(excelPath, excelName);
                        var data = ConvertExcel(excelReader, typeName);
                        var jsonPath = Path.Combine(Options.Default.JsonPath, $"{typeName}{Options.Default.JsonFileExtension}");
                        WriteToFile(data, jsonPath);
                    }
                } while (excelReader.NextResult());
            }
            catch (Exception e)
            {
                Console.WriteLine(
                    $"table <<<{Path.GetFileNameWithoutExtension(excelPath)}>>> parse json error ！！！！ \n{e}");
                throw;
            }
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private static object ConvertExcel(IExcelDataReader excelReader, string typeName)
        {
            List<int> ids = new List<int>();
            Dictionary<int, FieldInfo> column2Field = new Dictionary<int, FieldInfo>();
            var classType = RuntimeAssembly.RuntimeAsm.GetType(typeName);
            IList datas = Activator.CreateInstance(typeof(List<>).MakeGenericType(classType)) as IList;
            int row = 0;
            // 第一行是字段名 第二行是类型 第三行是注释
            while (excelReader.Read())
            {
                row++;
                if (row == 1)
                {
                    //第一行 记录一下字段
                    //这里重复判断是不是空，因为有的读取会读取到莫名的很多空数据
                    for (int i = 0; i < excelReader.FieldCount; i++)
                    {
                        var value = excelReader.GetString(i);
                        if (string.IsNullOrEmpty(value)) break;
                        value = value.TrimEnd();
                        value = value.TrimStart();
                        if (string.IsNullOrEmpty(value)) break;
                        if (value.StartsWith(Tools.IgnorePrefix))
                        {
                            continue;
                        }

                        var property = classType.GetField(value, BindingFlags.Instance | BindingFlags.Public);
                        if (property != null)
                            column2Field[i] = property;
                    }
                }

                if (row <= Tools.HeaderRows) continue;
                object targetClass = Activator.CreateInstance(classType);
                // row 是行 i是列
                foreach (var propertyInfo in column2Field)
                {
                    int column = propertyInfo.Key;
                    var strValue = excelReader.GetString(column);
                    if (string.IsNullOrEmpty(strValue)) strValue = String.Empty;
                    strValue = strValue.TrimEnd();
                    strValue = strValue.TrimStart();
                    //如果有一行的第一列是空，则直接跳出
                    if (column == 0 && (string.IsNullOrEmpty(strValue) || string.IsNullOrEmpty(strValue.Trim()))) break;
                    if (column == 0)
                    {
                        if (strValue == Tools.ExcelEndFlag) break;
                        if (strValue.StartsWith(Tools.IgnorePrefix)) continue;
                    }

                    object realValue = ParseStr(strValue, propertyInfo.Value.FieldType);
                    //默认第一列的id
                    if (column == 0)
                    {
                        if (ids.Contains((int)realValue))
                        {
                            throw new Exception($"{typeName} repeated id-{realValue}！！！！！");
                        }
                        ids.Add((int)realValue);
                        datas.Add(targetClass);
                    }

                    column2Field[column].SetValue(targetClass, realValue);
                }
            }
            var factoryType = RuntimeAssembly.RuntimeAsm.GetType(typeName + "Factory");
            object factory = Activator.CreateInstance(factoryType);
            var filedInfo = factoryType.GetField("list", BindingFlags.Public | BindingFlags.Instance);
            filedInfo.SetValue(factory, datas);
            return factory; 
        }

        private static void WriteToFile(object obj, string filePath)
        {
            if (Options.Default.IsBinaryJson == 1)
            {
                using FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write,
                    FileShare.ReadWrite);
                var data = obj.ToBson(serializer:new ObjectSerializer(type => true));
                file.Write(data, 0, data.Length);
            }
            else
            {
                
                var content = obj.ToJson(serializer:new ObjectSerializer(type => true));
                using FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write,
                    FileShare.ReadWrite); 
                using TextWriter writer = new StreamWriter(file, Options.Default.Encoding);
                writer.Write(content);
            }
        }

        private static object ParseStr(string strValue, Type propertyType)
        {
            if (string.IsNullOrEmpty(strValue)) strValue = String.Empty;
            strValue = strValue.TrimEnd();
            strValue = strValue.TrimStart();
            //如果有一行的第一列是空，则直接跳出
            object realValue = strValue;
            //根据类型不同填充属性
            //值为空
            if (string.IsNullOrEmpty(strValue) && (propertyType.IsValueType || propertyType == typeof(string)))
            {
                //如果是指类型则设置默认值
                if (propertyType.IsValueType)
                {
                    realValue = Activator.CreateInstance(propertyType);
                }

                if (propertyType == typeof(string))
                {
                    realValue = string.Empty;
                }
            }
            //如果是list
            else if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
            {
                IList list = (IList)Activator.CreateInstance(propertyType);
                if (!string.IsNullOrEmpty(strValue))
                {
                    var values = strValue.Split('|');

                    var genericType = propertyType.GenericTypeArguments[0];
                    foreach (string s in values)
                    {
                        list.Add(ParseStr(s, genericType));
                    }
                }

                realValue = list;
            }
            //如果是字典
            else if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Dictionary<,>))
            {
                IDictionary dic = (IDictionary)Activator.CreateInstance(propertyType);
                if (!string.IsNullOrEmpty(strValue))
                {
                    var values = strValue.Split('|');
                    var keyType = propertyType.GenericTypeArguments[0];
                    var valueType = propertyType.GenericTypeArguments[1];
                    foreach (var s in values)
                    {
                        var keyValues = s.Split(':');
                        var key = ParseStr(keyValues[0], keyType);
                        var val = ParseStr(keyValues[1], valueType);
                        dic.Add(key, val);
                    }
                }

                realValue = dic;
            }
            //如果是枚举
            else if (propertyType.IsEnum)
            {
                realValue = Enum.Parse(propertyType, strValue);
            }
            //基础类型
            else
            {
                realValue = Convert.ChangeType(strValue, propertyType);
            }

            return realValue;
        }
    }
}