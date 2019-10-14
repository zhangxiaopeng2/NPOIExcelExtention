using System.Collections.Generic;
using System.Data;

namespace NPOI.ExcelReaderHelper
{
    /// <summary>
    /// 表格转换为对象列表帮助类
    /// </summary>
    public static class DataTableToListHelper
    {

        /// <summary>
        /// DataTable中读取的Excel转换为对象列表
        /// </summary>
        /// <param name="dt">需要转换的Datable</param>
        /// <param name="jsonPath">键值路径</param>
        /// <returns></returns>
        public static List<T> DataTableToList<T>(this DataTable dt, string jsonPath) where T : class, new()
        {
            Dictionary<string, string> dic = JsonToDictionary.GetDicByJsonFile(jsonPath);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (dic.ContainsKey(dt.Columns[i].ColumnName))
                {
                    dt.Columns[i].ColumnName = dic[dt.Columns[i].ColumnName];
                }
            }
            var reflist = ModelConvertHelper.GetListByObject<T>(dt);
            return reflist;
        }

        /// <summary>
        /// 根据Excel路径转换为对应的对象
        /// </summary>
        /// <typeparam name="T">对象列表</typeparam>
        /// <param name="excelPath">excel文件路径</param>
        /// <param name="jsonPath">键值文件路径</param>
        /// <param name="sheetName">读取工作薄名字</param>
        /// <param name="firstLine">起始行，为内容中的上一行表头，从0开始</param>
        /// <returns></returns>
        public static List<T> DataTableToList<T>(this string excelPath, string jsonPath, string sheetName, int firstLine = 0) where T : class, new()
        {
            var reader = new ReaderHelper(excelPath);
            var dt = reader.ExcelToDataTable(sheetName, firstLine);
            Dictionary<string, string> dic = JsonToDictionary.GetDicByJsonFile(jsonPath);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (dic.ContainsKey(dt.Columns[i].ColumnName))
                {
                    dt.Columns[i].ColumnName = dic[dt.Columns[i].ColumnName];
                }
            }
            var reflist = ModelConvertHelper.GetListByObject<T>(dt);
            return reflist;
        }
    }
}
