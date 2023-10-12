using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Excel;
using ExcelToJson;

public class OneClassData
{
    public Dictionary<object, object> Data = new Dictionary<object, object>();
    public Dictionary<string, string> PropName2Comment = new Dictionary<string, string>();
    public string ClassName;
    public string FileName;
}

public static class DataFiller
{
    public static void FillData(string excelPath, Dictionary<string, OneClassData> datas)
    {
        try
        {
            using FileStream excelFile = File.Open(excelPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

            if (excelReader.ResultsCount < 1)
            {
                throw new Exception("Excel file is empty: " + excelPath);
            }

            do
            {
                var excelName = excelReader.Name;
                if (Tools.CheckTableIsEnum(excelName)) continue;
                if (Tools.CheckTableIsIgnore(excelName)) continue;
                var typeName = Tools.GetTableGenName(excelPath, excelName);
                var data = ReadExcel(excelReader, typeName, excelPath);
                if (datas.TryGetValue(typeName, out var oneClassData))
                {
                    oneClassData.Data = data;
                }
                else
                {
                    Console.WriteLine("compile class and not found data ??" + typeName);
                }
            } while (excelReader.NextResult());
        }
        catch (Exception e)
        {
            Console.WriteLine($"table {Path.GetFileNameWithoutExtension(excelPath)} parse json error ！！！！ \n{e}");
            throw;
        }
    }

    private static Dictionary<object, object> ReadExcel(IExcelDataReader excelReader, string typeName, string excelPath)
    {
        excelPath = Path.GetFileNameWithoutExtension(excelPath);
        Dictionary<object, object> importData = new Dictionary<object, object>();
        Dictionary<int, PropertyInfo> column2Property = new Dictionary<int, PropertyInfo>();
        var classType = RuntimeAssembly.RuntimeAsm.GetType(typeName);
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
                    if (i == 0)
                    {
                        if (value != "ID")
                            Console.WriteLine(
                                $"{Path.GetFileNameWithoutExtension(excelPath)}中{excelReader.Name}未标识为'ID'");
                    }

                    if (string.IsNullOrEmpty(value)) break;
                    value = value.TrimEnd();
                    value = value.TrimStart();
                    if (string.IsNullOrEmpty(value)) break;
                    if (value.StartsWith(Tools.IgnorePrefix))
                    {
                        continue;
                    }

                    var property = classType.GetProperty(value, BindingFlags.Instance | BindingFlags.Public);
                    if (property != null)
                        column2Property[i] = property;
                }
            }

            if (row <= Tools.HeaderRows)
            {
                continue;
            }

            object targetClass = Activator.CreateInstance(classType);
            // row 是行 i是列
            foreach (var propertyInfo in column2Property)
            {
                int column = propertyInfo.Key;
                var strValue = excelReader.GetString(column);
                if (string.IsNullOrEmpty(strValue)) strValue = String.Empty;
                strValue = strValue.TrimEnd();
                strValue = strValue.TrimStart();
                // 如果当前行第一列开头是#，则忽略
                if (column == 0 && strValue.StartsWith(Tools.IgnorePrefix)) continue;
                // 如果当前行第一列开头是结束标识，则结束表的生成
                if (column == 0 && strValue == Tools.ExcelEndFlag) return importData;
                //如果有一行的第一列是空，则直接跳出
                if (column == 0 && (string.IsNullOrEmpty(strValue) || string.IsNullOrEmpty(strValue.Trim()))) break;
                object realValue = strValue;
                var propertyType = column2Property[column].PropertyType;
                try
                {
                    realValue = GetValue(strValue, propertyType);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(
                        $"错误位于：{excelPath}表的{excelReader.Name}页签，{row}行{column}列的{propertyInfo.Value.Name}   {strValue}");
                    Console.WriteLine(ex.ToString());
                }

                //默认第一列的id
                if (column == 0)
                {
                    if (strValue == Tools.ExcelEndFlag) break;
                    if (strValue.StartsWith(Tools.IgnorePrefix)) continue;
                    if (importData.ContainsKey(realValue))
                    {
                        throw new Exception($"{typeName} repeated id-{realValue}！！！！！");
                    }

                    importData[realValue] = targetClass;
                }

                column2Property[column].SetValue(targetClass, realValue);
            }
        }

        return importData;
    }

    private static object GetValue(string data, Type type)
    {
        object outValue = null;
        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(List<>))
        {
            IList list = (IList)Activator.CreateInstance(type);
            if (!string.IsNullOrEmpty(data))
            {
                var values = data.Split('|');

                var genericType = type.GenericTypeArguments[0];
                foreach (string s in values)
                {
                    list.Add(GetSingleValue(s, genericType));
                }
            }

            outValue = list;
        }
        //如果是字典
        else if (type.IsGenericType &&
                 type.GetGenericTypeDefinition() == typeof(Dictionary<,>))
        {
            IDictionary dic = (IDictionary)Activator.CreateInstance(type);
            if (!string.IsNullOrEmpty(data))
            {
                var values = data.Split('|');
                var keyType = type.GenericTypeArguments[0];
                var valueType = type.GenericTypeArguments[1];
                foreach (var s in values)
                {
                    var keyValues = s.Split(':');
                    var key = GetSingleValue(keyValues[0], keyType);
                    var val = GetSingleValue(keyValues[1], valueType);
                    dic.Add(key, val);
                }
            }

            outValue = dic;
        }
        else
        {
            outValue = GetSingleValue(data, type);
        }

        return outValue;
    }

    private static object GetSingleValue(string data, Type type)
    {
        object outValue = null;
        //值为空
        if (string.IsNullOrEmpty(data))
        {
            //如果是指类型则设置默认值
            if (type.IsValueType)
            {
                outValue = Activator.CreateInstance(type);
            }
            else if (type == typeof(string))
            {
                outValue = string.Empty;
            }
        }
        //如果是枚举
        else if (type.IsEnum)
        {
            outValue = Enum.Parse(type, data);
        }
        else
        {
            outValue = Convert.ChangeType(data, type);
        }

        return outValue;
    }
}