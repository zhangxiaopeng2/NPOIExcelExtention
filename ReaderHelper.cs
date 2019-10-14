using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace NPOI.ExcelReaderHelper
{
    public class ReaderHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        /// <summary>
        /// 构造函数，保存的文件路径
        /// </summary>
        /// <param name="fileName"></param>
        public ReaderHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook();
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook();
            }

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// Excel转换为DataTable表格
        /// </summary>
        /// <param name="sheetName">Sheet名字,如果为空，则传null，默认读取第一个sheet</param>
        /// <param name="firstLine">起始行，其中为表头为起始行</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string sheetName, int firstLine = 0)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fs);
                }

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(firstLine);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        if (firstLine == 0)
                        {
                            DataColumn column = new DataColumn(firstRow.Cells[i].StringCellValue());
                            data.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(sheet.CellValue(firstLine, i));
                            data.Columns.Add(column);
                        }
                    }
                    startRow = firstLine + 1;

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null)
                        {
                            continue; //没有数据的行默认是null　　　　　　　
                        }

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            {
                                if (row.GetCell(j).CellType == CellType.Numeric)
                                {
                                    if (HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim().Replace(" ", "");
                                    }
                                }
                                else if (row.GetCell(j).CellType == CellType.Formula)//lxia 2017.12.04 begin
                                {
                                    try
                                    {
                                        if (HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                        {
                                            dataRow[j] = row.GetCell(j).DateCellValue.ToString("yyyy-MM-dd HH:mm:ss");
                                        }
                                        //不是日期型
                                        else
                                        {
                                            dataRow[j] = row.GetCell(j).NumericCellValue.ToString();
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim().Replace(" ", "");
                                    }
                                }////lxia 2017.12.04 end
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim().Replace(" ", "");
                                }
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                    {
                        fs.Close();
                    }
                }

                fs = null;
                disposed = true;
            }
        }

        private readonly int EXCEL03_MaxRow = 65535;

        /// <summary>
        /// 将DataTable转换为excel2003格式。
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public byte[] DataTable2Excel(DataTable dt, string sheetName)
        {

            IWorkbook book = new HSSFWorkbook();
            if (dt.Rows.Count < EXCEL03_MaxRow)
            {
                DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName);
            }
            else
            {
                int page = dt.Rows.Count / EXCEL03_MaxRow;
                for (int i = 0; i < page; i++)
                {
                    int start = i * EXCEL03_MaxRow;
                    int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                    DataWrite2Sheet(dt, start, end, book, sheetName + i.ToString());
                }
                int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
                DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString());
            }
            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// DataTable写入Sheet表格
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="startRow">起始行</param>
        /// <param name="endRow">终结行</param>
        /// <param name="book">工作单元</param>
        /// <param name="sheetName">工作薄名字</param>
        private void DataWrite2Sheet(DataTable dt, int startRow, int endRow, IWorkbook book, string sheetName)
        {
            ISheet sheet = book.CreateSheet(sheetName);

            IRow header = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = header.CreateCell(i);
                string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
                cell.SetCellValue(val);
            }
            int rowIndex = 1;
            for (int i = startRow; i <= endRow; i++)
            {
                DataRow dtRow = dt.Rows[i];
                IRow excelRow = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < dtRow.ItemArray.Length; j++)
                {
                    if (dtRow[j].GetType() == typeof(decimal))
                    {
                        ICell cell2 = excelRow.CreateCell(j, CellType.Numeric);
                        cell2.SetCellValue(Double.Parse(dtRow[j].ToString()));
                    }
                    else
                    {
                        excelRow.CreateCell(j).SetCellValue(dtRow[j].ToString());
                    }
                }
            }
            //sheet.ForceFormulaRecalculation = true; //设置导出的excel可使用公式
        }

        /// <summary>
        /// 判断字符串是否是int/double
        /// </summary>
        /// <param name="strNumber">需要判断的字符串</param>
        /// <returns></returns>
        public static bool IsDouble(string strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            const string strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            const string strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");
            return !objNotNumberPattern.IsMatch(strNumber) &&
                   !objTwoDotPattern.IsMatch(strNumber) &&
                   !objTwoMinusPattern.IsMatch(strNumber) &&
                   objNumberPattern.IsMatch(strNumber);
        }

        //将两个datatable合成一个
        public static DataTable UniteDataTable(DataTable dt1, DataTable dt2, string DTName)
        {
            DataTable dt3 = dt1.Clone();
            for (int i = 0; i < dt2.Columns.Count; i++)
            {
                dt3.Columns.Add(dt2.Columns[i].ColumnName);
            }
            object[] obj = new object[dt3.Columns.Count];

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                dt1.Rows[i].ItemArray.CopyTo(obj, 0);
                dt3.Rows.Add(obj);
            }

            if (dt1.Rows.Count >= dt2.Rows.Count)
            {
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    for (int j = 0; j < dt2.Columns.Count; j++)
                    {
                        dt3.Rows[i][j + dt1.Columns.Count] = dt2.Rows[i][j].ToString();
                    }
                }
            }
            else
            {
                DataRow dr3;
                for (int i = 0; i < dt2.Rows.Count - dt1.Rows.Count; i++)
                {
                    dr3 = dt3.NewRow();
                    dt3.Rows.Add(dr3);
                }
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    for (int j = 0; j < dt2.Columns.Count; j++)
                    {
                        dt3.Rows[i][j + dt1.Columns.Count] = dt2.Rows[i][j].ToString();
                    }
                }
            }
            dt3.TableName = DTName; //设置DT的名字
            return dt3;
        }

    }
}