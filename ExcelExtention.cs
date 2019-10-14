using System;
using System.Data;

namespace NPOI.ExcelReaderHelper
{
    public static class ExcelExtention
    {
        public static string GetValue(this DataRow[] row, string name)
        {
            try
            {
                return row[0][name].ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }
        public static string GetValue(this DataRow row, string name)
        {
            try
            {
                return row[name].ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
    public static class StringExtention
    {
        public static double GetValue(this string value)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return 0;
                }
                else
                {
                    return Convert.ToDouble(value);
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public static double GetValueFromPrent(this string value)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return 0;
                }
                else
                {
                    value = value.Replace("%", "");
                    return Convert.ToDouble(value);
                }
            }
            catch (Exception)
            {
                return 0;
            }

        }
    }
}

