# NPOIExcelExtention
NPOI读取Excel并转换为列表对象

# 描述
此帮助类为NPOI读取读取EXCEL的帮助类，其中.Net Framework版本为4.6，NPOI版本为2.4.1.0，通过NPOI读取EXCEL之后，使用反射的方式将对应的Datatable转换为相应的List，具体可参考附件中代码

# 调用示例
#region 单表头
var reader = new ReaderHelper("Import.xlsx");
var dt = reader.ExcelToDataTable(null);
var list = dt.DataTableToList<User>("JsonFile/user.txt");
Console.WriteLine($"单表头{JsonConvert.SerializeObject(list)}");
#endregion
#region 多表头
var readers = new ReaderHelper("Import1.xlsx");
var dts = readers.ExcelToDataTable(null, 1);
var list1 = dt.DataTableToList<User>("JsonFile/user.txt");
Console.WriteLine($"多表头{JsonConvert.SerializeObject(list1)}");
#endregion