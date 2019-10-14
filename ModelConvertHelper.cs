using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace NPOI.ExcelReaderHelper
{
    public class ModelConvertHelper
    {
        private static T RowConvertToObject<T>(DataRow hTable, T objModel)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(objModel);
            Type type = Type.GetType(objModel.GetType().AssemblyQualifiedName);
            foreach (PropertyDescriptor propertyDescriptor in properties)
            {
                PropertyInfo property = type.GetProperty(propertyDescriptor.Name);
                try
                {
                    if (hTable[propertyDescriptor.Name] != null)
                    {
                        if (property.PropertyType.FullName.ToString().ToLower().IndexOf("datetime") >= 0)
                        {
                            DateTime dateTime = DateTime.Parse(hTable[propertyDescriptor.Name].ToString());
                            property.SetValue(objModel, dateTime, null);
                        }
                        else
                        {
                            if (property.PropertyType.FullName.ToString().ToLower().IndexOf("decimal") >= 0)
                            {
                                decimal num = decimal.Parse(hTable[propertyDescriptor.Name].ToString());
                                property.SetValue(objModel, num, null);
                            }
                            else
                            {
                                if (property.PropertyType.FullName.ToString().ToLower().IndexOf("int64") >= 0)
                                {
                                    long num2 = long.Parse(hTable[propertyDescriptor.Name].ToString());
                                    property.SetValue(objModel, num2, null);
                                }
                                else
                                {
                                    if (property.PropertyType.FullName.ToString().ToLower().IndexOf("int") >= 0)
                                    {
                                        int num3 = int.Parse(hTable[propertyDescriptor.Name].ToString());
                                        property.SetValue(objModel, num3, null);
                                    }
                                    else
                                    {
                                        if (property.PropertyType.FullName.ToString().ToLower().IndexOf("long") >= 0)
                                        {
                                            long num4 = long.Parse(hTable[propertyDescriptor.Name].ToString());
                                            property.SetValue(objModel, num4, null);
                                        }
                                        else
                                        {
                                            object value = Convert.ChangeType(hTable[propertyDescriptor.Name], property.PropertyType);
                                            property.SetValue(objModel, value, null);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    property.SetValue(objModel, null, null);
                }
            }
            return objModel;
        }

        /// <summary>
        /// 根据键值给对应的属性赋值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<T> GetListByObject<T>(DataTable dt) where T : new()
        {
            List<T> list = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T obj = new T();
                list.Add(RowConvertToObject<T>(row, obj));
            }
            return list;
        }
        /// <summary>
        /// 对应的行数据转换成对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        public static T GetModel<T>(DataRow row) where T : new()
        {
            T obj = new T();
            return RowConvertToObject<T>(row, obj);
        }
    }
}
