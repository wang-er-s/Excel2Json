using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Framework;

public partial class SkinData : BaseConfig
{
    	/// id注释
	public Int32 ID{get; private set;}
	/// 名字
	public String Name{get; private set;}
	/// 测试列表
	public List<System.Int32> TestList{get; private set;}
	/// 测试枚举
	public List<ItemType> TestEnumList{get; private set;}
	/// 类型
	public ItemType Type{get; private set;}
	/// 字典
	public Dictionary<System.Int32,System.String> TestDic{get; private set;}
	/// 枚举字典
	public Dictionary<ItemType,System.Int32> TestEnumDic{get; private set;}
	/// 布尔
	public Boolean TestBool{get; private set;}
	/// 布尔
	public Boolean TestBool2{get; private set;}

}

[Config("#load_config_path")]
public partial class SkinDataFactory : ConfigSingleton<SkinDataFactory>
{
    private Dictionary<int, SkinData> dict = new Dictionary<int, SkinData>();

    [MongoDB.Bson.Serialization.Attributes.BsonElement]
    private List<SkinData> list = new List<SkinData>();

    public void Merge(SkinDataFactory o)
    {
        this.list.AddRange(o.list);
    }

    public override void EndInit()
    {
        foreach (SkinData config in list)
        {
            this.dict.Add(config.ID, config);
        }

        this.list.Clear();

        this.AfterEndInit();
    }
	
	partial void AfterEndInit();

    public SkinData Get(int id)
    {
        this.dict.TryGetValue(id, out SkinData SkinData);

        if (SkinData == null)
        {
            throw new Exception($"配置找不到，配置表名: {nameof(SkinData)}，配置id: {id}");
        }

        return SkinData;
    }

    public bool Contain(int id)
    {
        return this.dict.ContainsKey(id);
    }

    public IReadOnlyDictionary<int, SkinData> GetAll()
    {
        return this.dict;
    }

    public SkinData GetOne()
    {
        if (this.dict == null || this.dict.Count <= 0)
        {
            return null;
        }

        return list.GetRandomValue();
    }
}