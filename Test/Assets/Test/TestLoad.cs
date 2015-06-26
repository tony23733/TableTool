using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using TableTool;

public class TestLoad : MonoBehaviour
{
    /// <summary>
    /// 数据文件
    /// </summary>
    public TextAsset[] tables;
    /// <summary>
    /// 二进制数据字典
    /// </summary>
    private Dictionary<string, TextAsset> mTableList = new Dictionary<string, TextAsset>();

    // Use this for initialization
    void Awake()
    {
        Init();
    }

    void Init()
    {
        foreach (var v in tables)
            mTableList.Add(v.name, v);
        // TODO:在此处添加新的表
        LoadTableManager.FillTable<AreaMainKeyData>(mTableList["AreaMainKey"], typeof(AreaMainKeyManager));
        LoadTableManager.FillTable<AreaDoubleKeysData>(mTableList["AreaDoubleKeys"], typeof(AreaDoubleKeysManager));
    }

    void OnGUI()
    {
        GUILayout.BeginVertical();
        if (GUILayout.Button("print AreaMainKey"))
        {
            Debug.Log("--print AreaMainKey--");
            foreach (KeyValuePair<int, AreaMainKeyData> pair in AreaMainKeyManager.data)
            {
                Debug.Log("key = " + pair.Key + ", value:coordX = " + pair.Value.coordinate_x);
            }
        }
        if (GUILayout.Button("print AreaDoubleKeys"))
        {
            Debug.Log("--print AreaDoubleKeys--");
            foreach (KeyValuePair<AreaDoubleKeysKey, AreaDoubleKeysData> pair in AreaDoubleKeysManager.data)
            {
                Debug.Log("key:id = " + pair.Key.ID + ", value:coordX = " + pair.Value.coordinate_x);
            }
        }
        GUILayout.EndVertical();
    }
}
