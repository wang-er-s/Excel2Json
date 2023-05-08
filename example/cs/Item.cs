using System;
using System.Collections.Generic;
using Framework;

[Config("Item.bytes")]
public partial class Item
{
    /// <summary> id注释 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public int Id { get; private set; }
/// <summary> 名字 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public string Name { get; private set; }
/// <summary> 测试列表 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public List<int> TestList { get; private set; }
/// <summary> 测试枚举 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public List<ItemType> TestEnumList { get; private set; }
/// <summary> 类型 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public ItemType Type { get; private set; }
/// <summary> 字典 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	[MongoDB.Bson.Serialization.Attributes.BsonDictionaryOptions(MongoDB.Bson.Serialization.Options.DictionaryRepresentation.ArrayOfArrays)]
	public Dictionary<int, string> TestDic { get; private set; }
/// <summary> 枚举字典 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	[MongoDB.Bson.Serialization.Attributes.BsonDictionaryOptions(MongoDB.Bson.Serialization.Options.DictionaryRepresentation.ArrayOfArrays)]
	public Dictionary<ItemType, int> TestEnumDic { get; private set; }
/// <summary> 布尔 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public bool TestBool { get; private set; }
/// <summary> 布尔 </summary>
	[MongoDB.Bson.Serialization.Attributes.BsonElement]
	public bool TestBool2 { get; private set; }


    public static Dictionary<int, Item> Data { get; private set; }
	    
    private static void Load(string content)
    {
        if (Data != null) return;
        BeginInit();
        Data = SerializeHelper.DeSerialize<Dictionary<int, Item>>(content);
        EndInit();
    }

    public static Item GetById(int id)
    {
        if (Data.TryGetValue(id, out var result))
            return result;
        return null;
    }

    static partial void BeginInit();

    static partial void EndInit();
}