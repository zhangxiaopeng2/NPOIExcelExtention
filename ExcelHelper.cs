using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace NPOI.ExcelReaderHelper
{
    public static class ExcelHelper
    {
        public static CellRangeAddress GetRegion(this ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            if (!cell.IsMergedCell)
            {
                return new CellRangeAddress(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex);
            }

            for (int i = 0; i < cell.Sheet.NumMergedRegions; i++)
            {
                var mer = cell.Sheet.GetMergedRegion(i);
                if (mer.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return mer;
                }
            }
            return null;
        }
        public static ICell GetCell(this ISheet sheet, int? rownum, int? colnum)
        {
            if (sheet == null || rownum == null || colnum == null)
            {
                return null;
            }

            if (rownum < sheet.FirstRowNum
                || rownum > sheet.LastRowNum)
            {
                return null;
            }

            var row = sheet.GetRow(rownum.Value);
            if (row == null ||
                row.FirstCellNum > colnum ||
                row.LastCellNum < colnum)
            {
                return null;
            }

            var cell = row.GetCell(colnum.Value);
            return cell;
        }
        public static ICell GetFirstValueCell(this ISheet sheet, string value)
        {
            for (int i = sheet.FirstRowNum; i < sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                {
                    var cell = row.GetCell(j);
                    var cellvalue = cell?.StringCellValue();
                    if (cellvalue == null)
                    {
                        continue;
                    }

                    cellvalue = cellvalue.Replace("\r", "")
                        .Replace("\n", "")
                        .Replace("(", "")
                        .Replace(")", "")
                        .Replace(" ", "")
                        .Replace("（", "")
                        .Replace("）", "");
                    if (cellvalue.Trim() == value.Trim())
                    {
                        return cell;
                    }
                }
            }
            return null;
        }
        public static string GetRangeAddressValue(this ISheet sheet, CellRangeAddress cellrange,
            bool numRealValue = true, bool isdate = false)
        {
            if (cellrange == null)
            {
                return "";
            }

            return CellValue(sheet, cellrange.FirstRow, cellrange.FirstColumn, false, numRealValue, isdate);
        }
        /// <summary>
        /// 读取表格中单元格的数据
        /// </summary>
        /// <param name="sheet">表格</param>
        /// <param name="rownum">行号</param>
        /// <param name="colnum">列号</param>
        /// <param name="getMerged">是否查找合并的值</param>
        /// <param name="numRealValue">如果是数值型，是否返回表格中数值的真实数据，否则返回显示的有样式的数据</param>
        /// <param name="isdate">是否是时间格式</param>
        /// <returns></returns>
        public static string CellValue(this ISheet sheet, int rownum, int colnum,
            bool getMerged = true,
            bool numRealValue = true, bool isdate = false)
        {
            var cell = sheet.GetCell(rownum, colnum);
            if (cell == null)
            {
                return "";
            }

            if (cell.IsMergedCell && getMerged)
            {
                var re = cell.GetRegion();
                return CellValue(sheet, re.FirstRow, re.FirstColumn, false, numRealValue, isdate);
            }
            return cell.StringCellValue(numRealValue, isdate);
        }
        public static string StringCellValue(this ICell cell, bool numRealValue = true, bool isdate = false)
        {
            if (cell == null)
            {
                return null;
            }

            var celltype = cell.CellType;
            if (celltype == CellType.Formula)
            {
                celltype = cell.CachedFormulaResultType;
            }

            switch (celltype)
            {
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue?.Trim();
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString("yyyy/MM/dd");
                    }

                    if (numRealValue)
                    {
                        return cell.NumericCellValue.ToString()?.Trim();
                    }
                    else
                    {
                        ICellStyle cellstyle = cell.CellStyle as ICellStyle;
                        IDataFormat format = cell.Sheet.Workbook.CreateDataFormat();
                        var formatstring = format.GetFormat(cellstyle.DataFormat);
                        int length = formatstring.IndexOf('_');
                        if (length > 0)
                        {
                            formatstring = formatstring.Substring(0, length);
                        }

                        return cell.NumericCellValue.ToString(formatstring);
                    }
                case CellType.Error:
                    return cell.ErrorCellValue.ToString(CultureInfo.InvariantCulture);
                default:
                    return "";
            }
        }
        /// <summary>
        /// 创建一行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rownum"></param>
        /// <param name="rowHeightInPoints"></param>
        /// <returns></returns>
        public static IRow CreateRow(this ISheet sheet, int rownum, int rowHeightInPoints)
        {
            if (sheet == null)
            {
                return null;
            }

            var row = sheet.CreateRow(rownum);
            row.HeightInPoints = rowHeightInPoints;
            return row;
        }


        public static ICell CreateCell(this IRow row, int colnum, ICellStyle cellStyle)
        {
            if (row == null)
            {
                return null;
            }

            if (cellStyle == null)
            {
                cellStyle = row.Sheet.Workbook.CreateCellStyle();
                cellStyle.Alignment = HorizontalAlignment.Center;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BottomBorderColor = HSSFColor.Black.Index;
                cellStyle.BorderLeft = BorderStyle.Thin;
                cellStyle.LeftBorderColor = HSSFColor.Black.Index;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.RightBorderColor = HSSFColor.Black.Index;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.TopBorderColor = HSSFColor.Black.Index;
                var font = row.Sheet.Workbook.CreateFont();
                font.FontName = "华文细黑";
                cellStyle.SetFont(font);
            }
            var col = row.CreateCell(colnum);
            col.CellStyle = cellStyle;
            return col;
        }
        public static void SetColumnsWidth(this ISheet sheet, int firstcol, int lascolnum, int charwith)
        {
            if (sheet == null)
            {
                return;
            }

            for (int i = firstcol; i <= lascolnum; i++)
            {
                sheet.SetColumnWidth(i, 256 * charwith);
            }
        }


        /// <summary>
        /// 获取isheet
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public static ISheet GetSheet(this string filepath, string sheetname)
        {
            return OpenExcelFileAndOpenSheet(filepath, t => t.GetSheet(sheetname));
        }


        /// <summary>
        /// 获取isheet
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetindex"></param>
        /// <returns></returns>
        public static ISheet GetSheet(this string filepath, int sheetindex)
        {
            return OpenExcelFileAndOpenSheet(filepath, t => t.GetSheetAt(sheetindex));
        }


        /// <summary>
        /// 获取文件中的所有表的名称
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static string[] GetSheetNames(this IWorkbook workbook)
        {
            var result = new List<string>();
            var num = workbook.NumberOfSheets;
            for (int i = 0; i < num; i++)
            {
                result.Add(workbook.GetSheetAt(i).SheetName);
            }

            return result.ToArray();
        }


        public static IWorkbook GetWorkbook(this string filepath)
        {
            var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);

            try
            {
                if (Path.GetExtension(filepath)
                        ?.ToLower() == ".xls")
                {
                    var wk = new HSSFWorkbook(fs);
                    return wk;

                }
                else
                {
                    var wk = new XSSFWorkbook(fs);
                    return wk;
                }
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                fs.Dispose();
            }
        }


        private static ISheet OpenExcelFileAndOpenSheet(string filepath,
            Func<IWorkbook, ISheet> getfrombook)
        {
            var bk = GetWorkbook(filepath);

            if (bk == null)
            {
                return null;
            }

            return getfrombook(bk);
        }
    }
}
