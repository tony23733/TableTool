using System;
using System.IO;
using System.Collections.Generic;

namespace TableTool
{
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
}
