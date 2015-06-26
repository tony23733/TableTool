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
    public class OutputTable
    {
        /// <summary>
        /// 存储数据版本号
        /// </summary>
        public const int dataVersion = 1;

        /// <summary>
        /// 一键Build
        /// </summary>
        /// <param name="filePath"></param>
        public static void OneKeyBuild(string[] filePaths, CodeType codeTypes)
        {
            try
            {
                foreach (var v in filePaths)
                {
                    ExcelData excelData = ParseExcel(v);
                    if (excelData != null)
                    {
                        GenerateCodeData(excelData, codeTypes);
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
        public static void GenerateCodeData(ExcelData excelData, CodeType codeTypes)
        {
            List<IGenerateCode> list = new List<IGenerateCode>();
            if ((codeTypes & CodeType.U3D_CS) != CodeType.NULL)
            {
                list.Add(new U3dCsGenerateCode());
            }
            foreach (var v in list)
                v.GenerateCodeData(excelData);
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
            // 提取标记符号
            List<char> assignFlagSymbolList = new List<char>();     // 指定标记符号列表
            for (int i = 0; i < assignTypeSymbol.Length; ++i )
            {
                char c = assignTypeSymbol[i];
                bool b = false;
                switch(c)
                {
                    case '*': // 主键
                        data.mainKeyIndexs.Add(colmunsIndex);
                        break;
                    case '#': // 生成枚举定义代码
                        data.enumTypeIndexs.Add(colmunsIndex);
                        break;
                    case '-': // 忽略，不生成
                        data.ignoreFieldIndexs.Add(colmunsIndex);
                        break;
                    default:
                        b = true;
                        assignTypeSymbol = assignTypeSymbol.Substring(i);
                        break;
                }
                if (b)
                    break;
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

    [Flags]
    public enum CodeType
    {
        NULL = 0x0,
        U3D_CS = 0x1,
        JAVA = 0x2,
        CPP = 0x4,
    }
}
