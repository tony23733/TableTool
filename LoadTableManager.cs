using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System;
using System.Reflection;
using System.IO;

public class LoadTableManager : MonoBehaviour
{
    /// <summary>
    /// 表
    /// </summary>
    public TextAsset[] tables;
    private Dictionary<string, TextAsset> tableList = new Dictionary<string, TextAsset>();

    public static LoadTableManager instance { get; private set; }

	// Use this for initialization
	void Awake ()
	{
        if (instance != null)
            return;
        instance = this;
        DontDestroyOnLoad(this);
        Init();
	}

    void Init()
    {
        foreach (var v in tables)
            tableList.Add(v.name, v);
        // 在此处添加新的表
//         FillTable<AreaData>("Area", typeof(AreaManager));
//         FillTable<AttackpatternData>("Attackpattern", typeof(AttackpatternManager));
//         FillTable<BtnResData>("BtnRes", typeof(BtnResManager));
//         FillTable<MonsterDeckData>("MonsterDeck", typeof(MonsterDeckManager));
    }

    public static void FillTable<T>(string tableName, Type dataManagerType) where T : new ()
    {
        TextAsset tableAsset = instance.tableList[tableName];
        var list = LoadDataTemplate<T>(tableAsset);
        MethodInfo clearMethod = dataManagerType.GetMethod("Clear");
        clearMethod.Invoke(null, null);
        MethodInfo addItemMethod = dataManagerType.GetMethod("AddItem");
        foreach (var v in list)
        {
            addItemMethod.Invoke(null, new object[] { v });
        }
    }

    public static List<T> LoadDataTemplate<T>(TextAsset data) where T : new()
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
    }
}
