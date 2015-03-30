/*
 * 生成时间：2015/3/30 13:19//例子：2014/11/20 17:57
 * 表格文件：AreaMainKey.xls//例子：Area.xls
 * 检索键：ID //例子：Name Type
 */
using UnityEngine;
using System.Collections;
using System.Collections.Generic;

public struct AreaMainKeyData
{
    /// <summary>
    /// 区域ID
    /// 多行测试
    /// </summary>
    public int ID;
    /// <summary>
    /// 类型
    /// </summary>
    public typeEnum type;
    /// <summary>
    /// 描述
    /// </summary>
    public string Des;
    /// <summary>
    /// 中心格 X
    /// </summary>
    public int coordinate_x;
    /// <summary>
    /// 中心格 Y
    /// </summary>
    public float coordinate_y;
    /// <summary>
    /// 参数1
    /// </summary>
    public long param1;
    /// <summary>
    /// 参数2
    /// </summary>
    public double param2;
    /// <summary>
    /// 获取key
    /// </summary>
    public int GetKey()
    {
		var key = ID;
        return key;
    }
}

public enum typeEnum
{
	Empty=0,
	All=1,
	Half=2
}

public class AreaMainKeyManager
{
    private static Dictionary<int, AreaMainKeyData> mData = new Dictionary<int, AreaMainKeyData>();
    /// <summary>
    /// 表数据
    /// </summary>
    public static Dictionary<int, AreaMainKeyData> data { get { return mData; } }
    /// <summary>
    /// 清除所有数据
    /// </summary>
    public static void Clear()
    {
        mData.Clear();
    }
    /// <summary>
    /// 添加成员
    /// </summary>
    public static void AddItem(AreaMainKeyData item)
    {
        var key = item.GetKey();
        mData.Add(key, item);
    }
    /// <summary>
    /// 获取数据
    /// </summary>
    public static AreaMainKeyData GetItem(int ID)
    {
		var key = ID;
        return mData[key];
    }
}