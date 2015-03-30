using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows;
// using Microsoft.Office.Core;
// using Microsoft.Office.Interop.Excel;
using System.IO;
// using System.Reflection;

namespace TableTool
{
    class OutputTable
    {
        /// <summary>
        /// 存储数据版本号
        /// </summary>
        public const int dataVersion = 1;
        /// <summary>
        /// 代码头模板
        /// </summary>
        public const string codeHeadTemplate =
@"/*
 * 生成时间：{0}//例子：2014/11/20 17:57
 * 表格文件：{1}//例子：Area.xls
 * 检索键：{2}//例子：Name Type
 */
using UnityEngine;
using System.Collections;
using System.Collections.Generic;
";
        /// <summary>
        /// 代码结构体模板。"/[","/]"分别代表"{","}"
        /// </summary>
        public const string codeStructTemplate =
@"public struct {0}
/[
{1}
/]
";
        /// <summary>
        /// 代码获取key模板
        /// </summary>
        public const string codeGetKeyTemplate =
@"    /// <summary>
    /// 获取key
    /// </summary>
    public {0} GetKey()
    /[
{1}
        return key;
    /]";
        /// <summary>
        /// 代码变量模板
        /// </summary>
        public const string codeVariableTemplate = 
@"    /// <summary>
    /// {0}
    /// </summary>
    public {1} {2};";
        /// <summary>
        /// 代码枚举模板
        /// </summary>
        public const string codeEnumTemplate =
@"public enum {0}
/[
{1}
/]
";
        /// <summary>
        /// 代码管理器模板
        /// </summary>
        public const string codeManagerTemplate =
@"public class {0}Manager
/[
    private static Dictionary<{1}, {2}> mData = new Dictionary<{3}, {4}>();
    /// <summary>
    /// 表数据
    /// </summary>
    public static Dictionary<{5}, {6}> data /[ get /[ return mData; /] /]
    /// <summary>
    /// 清除所有数据
    /// </summary>
    public static void Clear()
    /[
        mData.Clear();
    /]
    /// <summary>
    /// 添加成员
    /// </summary>
    public static void AddItem({7} item)
    /[
        var key = item.GetKey();
        mData.Add(key, item);
    /]
    /// <summary>
    /// 获取数据
    /// </summary>
    public static {8} GetItem({9})
    /[
{10}
        return mData[key];
    /]
/]";
        /// <summary>
        /// 一键Build
        /// </summary>
        /// <param name="filePath"></param>
        public static void OneKeyBuild(string[] filePaths)
        {
            try
            {
                foreach (var v in filePaths)
                {
                    ExcelData excelData = ParseExcel(v);
                    if (excelData != null)
                    {
                        GenerateCodeData(excelData);
                        GenerateBinaryData(excelData);
                    }
                }
                MessageBox.Show("生成完成！", "Build");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Build");
            }
        }

        /// <summary>
        /// 加载excel表数据
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataSet LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
//                 FileInfo fileInfo = new FileInfo(filePath);
//                 if (fileInfo.Extension == ".xls")
//                     strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
//                 else
//                     strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                DataTable dt = OleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                List<string> sheetNames = new List<string>();
                foreach (DataRow row in dt.Rows)
                    sheetNames.Add(row["TABLE_NAME"].ToString());
                String sql = "SELECT * FROM  [" + sheetNames[0] + "]";      // 格式：["sheet1$"]，sheetNames末尾含$符
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, sheetNames[0]);
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "打开失败");
                return null;
            }
        }

        /// <summary>
        /// 解析excel表
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static ExcelData ParseExcel(string filePath)
        {
            DataSet ds = LoadDataFromExcel(filePath);
            if (ds == null)
                return null;
            ExcelData data = new ExcelData();
            data.fullFileName = filePath;
            try
            {
                DataTable table = ds.Tables[0];
                DataColumn[] columns = new DataColumn[table.Columns.Count];
                table.Columns.CopyTo(columns, 0);
                foreach (var v in columns)      // columns是第一行数据
                {
                    data.names.Add(v.ToString().Trim().Replace(' ', '_'));      // 去掉收尾空字符，将中间空格改为下划线
                }
                int row = 0;
                foreach (DataRow dr in table.Rows)
                {
                    for (int c = 0; c < columns.Length; ++c )
                    {
                        object obj = dr[columns[c]];
                        if (row == 0)       // 类型，实际是第二行数据
                        {
                            ParseType(data, obj, c);
                            continue;
                        }
                        if (row == 1)       // 说明，用于代码注释
                        {
                            data.descriptions.Add(obj.ToString());
                            continue;
                        }
                        int valueIndex = row - 2;       // row从2开始
                        if (data.values.Count < row - 1)
                            data.values.Add(new List<string>());
                        data.values[valueIndex].Add(obj.ToString());
                    }
                    ++row;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "数据错误");
                return null;
            }
            return data;
        }

        /// <summary>
        /// 生成代码
        /// </summary>
        /// <param name="excelData"></param>
        public static void GenerateCodeData(ExcelData excelData)
        {
            StringBuilder codeBuilder = new StringBuilder();
            // 生成文件头
            string date = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            StringBuilder mainKeyBuilder = new StringBuilder();
            foreach (var v in excelData.mainKeyIndexs)
                mainKeyBuilder.Append(excelData.names[v]).Append(" ");
            string headText = string.Format(codeHeadTemplate, date, excelData.fileNameWithExtension, mainKeyBuilder);
            codeBuilder.Append(headText);
            // 生成主键，key超过1时生成数据结构
            if (excelData.mainKeyIndexs.Count > 1)
            {
                TableDataType[] mainKeyTypes = new TableDataType[excelData.mainKeyIndexs.Count];
                string[] mainKeyAttributeNames = new string[excelData.mainKeyIndexs.Count];
                string[] mainKeyComments = new string[excelData.mainKeyIndexs.Count];
                for (int i = 0; i < excelData.mainKeyIndexs.Count; ++i)
                {
                    int index = excelData.mainKeyIndexs[i];
                    mainKeyTypes[i] = excelData.types[index];
                    mainKeyAttributeNames[i] = excelData.names[index];
                    mainKeyComments[i] = excelData.descriptions[index];
                }
                string mainKeyAttributeCode = GenerateAttributeCode(mainKeyTypes, mainKeyAttributeNames, mainKeyComments);
                codeBuilder.Append("\r\n").Append(FormatCode(
                    codeStructTemplate, GetKeyStructName(excelData.fileName), mainKeyAttributeCode
                    ));
            }
            // 生成数据结构
            StringBuilder dataCodeBuilder = new StringBuilder();        // 数据代码
            string dataAttributeCode = GenerateAttributeCode(excelData.types.ToArray(), excelData.names.ToArray(), excelData.descriptions.ToArray());
            dataCodeBuilder.Append(dataAttributeCode);
            StringBuilder newKeyBuilder = new StringBuilder();      // 实例化key
            string keyTypeString = "";
            if (excelData.mainKeyIndexs.Count > 1)      // 多个参数时
            {
                newKeyBuilder.Append("\t\tvar key = new ").Append(GetKeyStructName(excelData.fileName)).Append("();\r\n");
                for (int i = 0; i < excelData.mainKeyIndexs.Count; ++i)
                {
                    int index = excelData.mainKeyIndexs[i];
                    string paramName = excelData.names[index];
                    newKeyBuilder.Append("\t\tkey.").Append(paramName).Append(" = ").Append(paramName).Append(";");
                    if (i != excelData.mainKeyIndexs.Count - 1)
                        newKeyBuilder.Append("\r\n");
                }
                keyTypeString = GetKeyStructName(excelData.fileName);
            }
            else
            {
                int index = excelData.mainKeyIndexs[0];
                newKeyBuilder.Append("\t\tvar key = ").Append(excelData.names[index]).Append(";");
                keyTypeString = GenerateTypeByTableDataType(excelData.types[index], excelData.names[index]);
            }
            dataCodeBuilder.Append("\r\n").Append(FormatCode(codeGetKeyTemplate, keyTypeString, newKeyBuilder));
            codeBuilder.Append("\r\n").Append(FormatCode(
                codeStructTemplate, GetDataStructName(excelData.fileName), dataCodeBuilder
                ));
            // 生成枚举
            List<int> enumIndexs = new List<int>();
            for (int i = 0; i < excelData.types.Count; ++i)
            {
                if (excelData.types[i] == TableDataType.ENUM)
                {
                    enumIndexs.Add(i);
                }
            }
            string[] attributeNames = new string[enumIndexs.Count];
            string[] typeStrings = new string[enumIndexs.Count];
            for (int i = 0; i < enumIndexs.Count; ++i)
            {
                int index = enumIndexs[i];
                attributeNames[i] = excelData.names[index];
                typeStrings[i] = excelData.typeStrings[index];
            }
            string enumCode = GenerateEnemCode(attributeNames, typeStrings);
            codeBuilder.Append(enumCode);
            // 生成数据管理器
            string[] managerCodeUnits = new string[11];
            managerCodeUnits[0] = excelData.fileName;
            string keyType = "";
            if (excelData.mainKeyIndexs.Count > 1)
            {
                keyType = GetKeyStructName(excelData.fileName);
            }
            else
            {
                int index = excelData.mainKeyIndexs[0];
                keyType = GenerateTypeByTableDataType(excelData.types[index], excelData.names[index]);
            }
            managerCodeUnits[1] = managerCodeUnits[3] = managerCodeUnits[5] = keyType;
            managerCodeUnits[2] = managerCodeUnits[4] = managerCodeUnits[6] = managerCodeUnits[7] =
                managerCodeUnits[8] = GetDataStructName(excelData.fileName);
            StringBuilder addItemParamBuilder = new StringBuilder();
            for (int i = 0; i < excelData.mainKeyIndexs.Count; ++i)
            {
                int index = excelData.mainKeyIndexs[i];
                string paramType = GenerateTypeByTableDataType(excelData.types[index], excelData.names[i]);
                addItemParamBuilder.Append(paramType).Append(" ").Append(excelData.names[i]);
                if (i != excelData.mainKeyIndexs.Count - 1)
                    addItemParamBuilder.Append(", ");
            }
            managerCodeUnits[9] = addItemParamBuilder.ToString();
            managerCodeUnits[10] = newKeyBuilder.ToString();     // 与GetKey函数相同
            codeBuilder.Append("\r\n").Append(FormatCode(codeManagerTemplate, managerCodeUnits));
            // 保存代码文件
            string codeFilePath = excelData.path + "/" + excelData.fileName + ".cs";
            using (StreamWriter sw = new StreamWriter(codeFilePath, false))
            {
                sw.Write(codeBuilder);
            }
        }

        /// <summary>
        /// 生成二进制数据
        /// </summary>
        /// <param name="excelData"></param>
        public static void GenerateBinaryData(ExcelData excelData)
        {
            string bytesFileName = excelData.path + "/" + excelData.fileName + ".bytes";
            using (BinaryWriter writer = new BinaryWriter(File.Open(bytesFileName, FileMode.Create)))
            {
                writer.Write(dataVersion);      // 版本号
                writer.Write(excelData.values.Count);
                foreach (var v in excelData.types)
                    writer.Write((Byte)v);
                foreach (var lineData in excelData.values)
                {
                    for (int i = 0; i < lineData.Count; ++i)
                    {
                        switch (excelData.types[i])
                        {
                            case TableDataType.INT:
                                int intData = 0;
                                int.TryParse(lineData[i], out intData);
                                writer.Write(intData);
                                break;
                            case TableDataType.FLOAT:
                                float floatData = 0f;
                                float.TryParse(lineData[i], out floatData);
                                writer.Write(floatData);
                                break;
                            case TableDataType.STRING:
                                writer.Write(lineData[i]);
                                break;
                            case TableDataType.ENUM:
                                int enumData = 0;
                                int.TryParse(lineData[i], out enumData);
                                writer.Write(enumData);
                                break;
                            case TableDataType.LONG:
                                long longData = 0L;
                                long.TryParse(lineData[i], out longData);
                                writer.Write(longData);
                                break;
                            case TableDataType.DOUBLE:
                                double doubleData = 0d;
                                double.TryParse(lineData[i], out doubleData);
                                writer.Write(doubleData);
                                break;
                            // TODO:根据需求扩展类型
                            default:
                                break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 解析指定类型
        /// </summary>
        /// <param name="data"></param>
        /// <param name="obj"></param>
        /// <param name="colmunsIndex"></param>
        private static void ParseType(ExcelData data, object obj, int colmunsIndex)
        {
            string assignTypeString = obj.ToString().Trim();
            data.typeStrings.Add(assignTypeString);
            string assignTypeSymbol = assignTypeString.Split(new char[] { '\n', '\r' }, System.StringSplitOptions.RemoveEmptyEntries)[0];
            if (assignTypeSymbol[0] == '*')     // 主键
            {
                assignTypeSymbol = assignTypeSymbol.Substring(1);
                data.mainKeyIndexs.Add(colmunsIndex);
            }
            switch (assignTypeSymbol)
            {
                case "INT":
                    data.types.Add(TableDataType.INT);
                    break;
                case "FLOAT":
                    data.types.Add(TableDataType.FLOAT);
                    break;
                case "STRING":
                    data.types.Add(TableDataType.STRING);
                    break;
                case "ENUM":
                    data.types.Add(TableDataType.ENUM);
                    break;
                case "LONG":
                    data.types.Add(TableDataType.LONG);
                    break;
                case "DOUBLE":
                    data.types.Add(TableDataType.DOUBLE);
                    break;
                // TODO:根据需求扩展类型
                default:
                    Console.WriteLine("指定类型{0}不存在", assignTypeSymbol);
                    break;
            }
        }

        /// <summary>
        /// 获取枚举类型名
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetEnumTypeName(string name)
        {
            return name + "Enum";
        }

        /// <summary>
        /// 获取key结构体名，有多个key时使用
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetKeyStructName(string name)
        {
            return name + "Key";
        }

        /// <summary>
        /// 获取数据结构体名
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetDataStructName(string name)
        {
            return name + "Data";
        }

        /// <summary>
        /// 生成基本类型字符
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static string GenerateTypeByTableDataType(TableDataType type, string attributeName)
        {
            switch (type)
            {
                case TableDataType.INT:
                    return "int";
                case TableDataType.FLOAT:
                    return "float";
                case TableDataType.STRING:
                    return "string";
                case TableDataType.ENUM:
                    return GetEnumTypeName(attributeName);
                case TableDataType.LONG:
                    return "long";
                case TableDataType.DOUBLE:
                    return "double";
                // TODO:根据需求扩展类型
                default:
                    return "";
            }
        }

        /// <summary>
        /// 生成属性代码
        /// </summary>
        /// <param name="type"></param>
        /// <param name="attributeName"></param>
        /// <param name="comment"></param>
        /// <returns></returns>
        private static string GenerateAttributeCode(TableDataType type, string attributeName, string comment)
        {
            string typeName = GenerateTypeByTableDataType(type, attributeName);
            if (string.IsNullOrEmpty(typeName))
                return "";
            string[] commentSplit = comment.Split(new char[] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries);
            StringBuilder commentBuilder = new StringBuilder();
            for (int i = 0; i < commentSplit.Length; ++i)
            {
                commentBuilder.Append(commentSplit[i]);
                if (i != commentSplit.Length - 1)
                    commentBuilder.Append(@"
    /// ");
            }
            return string.Format(codeVariableTemplate, commentBuilder, typeName, attributeName);
        }

        /// <summary>
        /// 生成属性代码
        /// </summary>
        /// <param name="types"></param>
        /// <param name="attributeNames"></param>
        /// <param name="comments"></param>
        /// <returns></returns>
        private static string GenerateAttributeCode(TableDataType[] types, string[] attributeNames, string[] comments)
        {
            if (types.Length != attributeNames.Length)
            {
                Console.WriteLine("类型数量与变量数量不匹配。");
                return "";
            }
            StringBuilder attributeBuilder = new StringBuilder();
            for (int i = 0; i < types.Length; ++i)
            {
                attributeBuilder.Append(GenerateAttributeCode(types[i], attributeNames[i], comments[i]));
                if (i != types.Length - 1)
                    attributeBuilder.Append("\r\n");
            }
            return attributeBuilder.ToString();
        }

        /// <summary>
        /// 生成枚举代码
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="typeString"></param>
        /// <returns></returns>
        private static string GenerateEnemCode(string attributeName, string typeString)
        {
            StringBuilder enumValueBuilder = new StringBuilder();
            string[] typeStringSplit = typeString.Split(new char[] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries);
            for (int i = 1; i < typeStringSplit.Length; ++i)        // 第一行是类型，后面行是枚举值
            {
                enumValueBuilder.Append("\t").Append(typeStringSplit[i]);
                if (i != typeStringSplit.Length - 1)
                    enumValueBuilder.Append(",\r\n");
            }
            return FormatCode(codeEnumTemplate, GetEnumTypeName(attributeName), enumValueBuilder);
        }

        /// <summary>
        /// 生成枚举代码
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="typeString"></param>
        /// <returns></returns>
        private static string GenerateEnemCode(string[] attributeNames, string[] typeStrings)
        {
            if (attributeNames.Length != typeStrings.Length)
            {
                Console.WriteLine("变量数量与类型数量不匹配。");
                return "";
            }
            StringBuilder enumBuilder = new StringBuilder("\r\n");      // 先换行
            for (int i = 0; i < attributeNames.Length; ++i)
            {
                enumBuilder.Append(GenerateEnemCode(attributeNames[i], typeStrings[i]));
                if (i != attributeNames.Length - 1)
                    enumBuilder.Append("\r\n");
            }
            return enumBuilder.ToString();
        }

        /// <summary>
        /// 格式化代码性文本，会将\[转换为{
        /// </summary>
        /// <param name="codeTemplate"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private static string FormatCode(string codeTemplate, params object[] args)
        {
            string code = string.Format(codeTemplate, args);
            code = code.Replace("/[", "{").Replace("/]", "}");
            return code;
        }
    }

    /// <summary>
    /// excel表数据
    /// </summary>
    public class ExcelData
    {
        /// <summary>
        /// excel表完整路径名
        /// </summary>
        public string fullFileName;
        /// <summary>
        /// excel表文件名，无路径无扩展名
        /// </summary>
        public string fileName { get { return Path.GetFileNameWithoutExtension(fullFileName); } }
        /// <summary>
        /// excel表文件名，无路径有扩展名
        /// </summary>
        public string fileNameWithExtension { get { return Path.GetFileName(fullFileName); } }
        /// <summary>
        /// excel表文件路径，也是生成路径
        /// </summary>
        public string path { get { return Path.GetDirectoryName(fullFileName); } }
        /// <summary>
        /// 转换后的变量名
        /// </summary>
        public List<string> names = new List<string>();
        /// <summary>
        /// 转换后的类型
        /// </summary>
        public List<TableDataType> types = new List<TableDataType>();
        /// <summary>
        /// 原本类型描述文本
        /// </summary>
        public List<string> typeStrings = new List<string>();
        /// <summary>
        /// 中文说明，用作代码注释
        /// </summary>
        public List<string> descriptions = new List<string>();
        /// <summary>
        /// 二维值列表文本
        /// </summary>
        public List<List<string>> values = new List<List<string>>();
        /// <summary>
        /// 关键key索引
        /// </summary>
        public List<int> mainKeyIndexs = new List<int>();
    }

    /// <summary>
    /// 表数据类型，与数据读取定义的格式需要同步
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
