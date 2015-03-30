using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System;
using System.Reflection;
using System.IO;

public class LoadTableManager
{
	/// <summary>
	/// 解析填充数据
	/// </summary>
	/// <param name="tableAsset">二进制表数据</param>
	/// <param name="dataManagerType">自动生成的数据管理器类型</param>
	/// <typeparam name="T">自动生成的数据类型</typeparam>
	public static void FillTable<T>(TextAsset tableAsset, Type dataManagerType) where T : new ()
    {
        var list = LoadDataTemplate<T>(tableAsset);
        MethodInfo clearMethod = dataManagerType.GetMethod("Clear");
        clearMethod.Invoke(null, null);
        MethodInfo addItemMethod = dataManagerType.GetMethod("AddItem");
        foreach (var v in list)
        {
            addItemMethod.Invoke(null, new object[] { v });
        }
    }
	
	/// <summary>
	/// 加载数据模板
	/// </summary>
	/// <returns>数据列表</returns>
	/// <param name="data">二进制表数据</param>
	/// <typeparam name="T">自动生成的数据类型</typeparam>
    private static List<T> LoadDataTemplate<T>(TextAsset data) where T : new()
    {
        List<T> list = new List<T>();
        using (BinaryReader reader = new BinaryReader(new MemoryStream(data.bytes)))
        {
            Type type = typeof(T);
            FieldInfo[] fields = type.GetFields();
            TableDataType[] dataTypes = new TableDataType[fields.Length];
            /*int version = */reader.ReadInt32();
            int dataLength = reader.ReadInt32();
            for (int i = 0; i < dataTypes.Length;++i)
            {
                dataTypes[i] = (TableDataType)reader.ReadByte();
            }
            for (int dataIndex = 0; dataIndex < dataLength; ++dataIndex)
            {
                T dataT = new T();
                object dataObj = dataT as object;
                for (int fieldIndex = 0; fieldIndex < fields.Length; ++fieldIndex)
                {
                    object fieldData = null;
                    switch (dataTypes[fieldIndex])
                    {
                        case TableDataType.INT:
                            fieldData = reader.ReadInt32();
                            break;
                        case TableDataType.FLOAT:
                            fieldData = reader.ReadSingle();
                            break;
                        case TableDataType.STRING:
                            fieldData = reader.ReadString();
                            break;
                        case TableDataType.ENUM:
                            fieldData = reader.ReadInt32();
                            break;
                        case TableDataType.LONG:
                            fieldData = reader.ReadInt64();
                            break;
                        case TableDataType.DOUBLE:
                            fieldData = reader.ReadDouble();
                            break;
						// TODO:根据需要扩展类型
                        default: break;
                    }
                    if (fieldData != null)
                    {
                        fields[fieldIndex].SetValue(dataObj, fieldData);
                        dataT = (T)dataObj;
                    }
                    else
                    {
                        Debug.LogWarning("--Toto-- LoadTableManager->LoadDataTemplate: fieldData is null.");
                        continue;
                    }
                }
                list.Add(dataT);
            }
        }

        return list;
    }

    /// <summary>
    /// 表数据类型，与代码生成器定义的格式需要同步
    /// </summary>
    public enum TableDataType
    {
        NULL = 0,
        INT,
        FLOAT,
        STRING,
        ENUM,
        LONG,
        DOUBLE
		// TODO:根据需求扩展类型
    }
}
