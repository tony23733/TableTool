using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace TableTool
{
    public class U3dCsGenerateCode : IGenerateCode
    {
        /// <summary>
        /// 代码头模板
        /// eg.生成时间：2015/6/26 9:24
        /// 表格文件：AreaMainKey.xls
        /// 检索键：ID
        /// </summary>
        public const string codeHeadTemplate =
@"/*
 * 生成时间：{0}
 * 表格文件：{1}
 * 检索关键字段：{2}
 * 忽略字段：{3}
 */
using System.Collections;
using System.Collections.Generic;
";
        public const string namespaceTemplate =
@"namespace TableTool
/[
{0}
/]
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

        public void GenerateCodeData(ExcelData excelData)
        {
            StringBuilder codeBuilder = new StringBuilder();
            // 生成文件头
            string date = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            StringBuilder mainKeyBuilder = new StringBuilder();     // 文件头关键字段说明
            foreach (var v in excelData.mainKeyIndexs)
                mainKeyBuilder.Append(excelData.names[v]).Append(" ");
            StringBuilder ignoreKeyBuilder = new StringBuilder();       // 文件头忽略字段说明
            foreach (var v in excelData.ignoreFieldIndexs)
                ignoreKeyBuilder.Append(excelData.names[v]).Append(" ");
            string headText = FormatCode(codeHeadTemplate, date, excelData.fileNameWithExtension, mainKeyBuilder, ignoreKeyBuilder);
            codeBuilder.Append(headText);
            StringBuilder classCodeBuilder = new StringBuilder();       // 类代码部分，不含文件头说明和命名空间
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
                classCodeBuilder.Append(FormatCode(
                    codeStructTemplate, GetKeyStructName(excelData.fileName), mainKeyAttributeCode
                    ));
            }
            // 生成数据结构
            StringBuilder dataCodeBuilder = new StringBuilder();        // 数据代码
            string dataAttributeCode = GenerateAttributeCode(excelData.types.ToArray(), excelData.names.ToArray(), excelData.descriptions.ToArray());
            dataCodeBuilder.Append(dataAttributeCode).Append("\r\n");
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
            dataCodeBuilder.Append(FormatCode(codeGetKeyTemplate, keyTypeString, newKeyBuilder));
            classCodeBuilder.Append(FormatCode(
                codeStructTemplate, GetDataStructName(excelData.fileName), dataCodeBuilder
                ));
            // 生成枚举
            List<int> enumIndexs = new List<int>();
            for (int i = 0; i < excelData.types.Count; ++i)
            {
                if (excelData.types[i] == TableDataType.ENUM && excelData.enumTypeIndexs.Contains(i))       // 判断是枚举类型且标记"#"
                {
                    enumIndexs.Add(i);
                }
            }
            if (enumIndexs.Count > 0)
            {
                string[] attributeNames = new string[enumIndexs.Count];
                string[] typeStrings = new string[enumIndexs.Count];
                for (int i = 0; i < enumIndexs.Count; ++i)
                {
                    int index = enumIndexs[i];
                    attributeNames[i] = excelData.names[index];
                    typeStrings[i] = excelData.typeStrings[index];
                }
                string enumCode = GenerateEnemCode(attributeNames, typeStrings);
                classCodeBuilder.Append(enumCode);
            }
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
            classCodeBuilder.Append(FormatCode(codeManagerTemplate, managerCodeUnits));
            // 将类代码、命名空间添加到代码
            string namespaceClassCode = FormatCode(namespaceTemplate, classCodeBuilder);
            codeBuilder.Append(namespaceClassCode);
            // 保存代码文件
            string codeFilePath = excelData.path + "/" + excelData.fileName + ".cs";
            using (StreamWriter sw = new StreamWriter(codeFilePath, false))
            {
                sw.Write(codeBuilder);
            }
        }

        /// <summary>
        /// 获取枚举类型名
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string GetEnumTypeName(string name)
        {
            return name + "Enum";
        }

        /// <summary>
        /// 获取key结构体名，有多个key时使用
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string GetKeyStructName(string name)
        {
            return name + "Key";
        }

        /// <summary>
        /// 获取数据结构体名
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string GetDataStructName(string name)
        {
            return name + "Data";
        }

        /// <summary>
        /// 生成基本类型字符
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private string GenerateTypeByTableDataType(TableDataType type, string attributeName)
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
        private string GenerateAttributeCode(TableDataType type, string attributeName, string comment)
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
        private string GenerateAttributeCode(TableDataType[] types, string[] attributeNames, string[] comments)
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
        private string GenerateEnemCode(string attributeName, string typeString)
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
        private string GenerateEnemCode(string[] attributeNames, string[] typeStrings)
        {
            if (attributeNames.Length != typeStrings.Length)
            {
                Console.WriteLine("变量数量与类型数量不匹配。");
                return "";
            }
            StringBuilder enumBuilder = new StringBuilder();
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
        private string FormatCode(string codeTemplate, params object[] args)
        {
            string code = string.Format(codeTemplate, args);
            code = code.Replace("/[", "{").Replace("/]", "}");
            return code;
        }
    }
}
