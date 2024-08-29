// See https://aka.ms/new-console-template for more information

using Newtonsoft.Json;
using OfficeOpenXml;
using System.Reflection;
using Formatting = Newtonsoft.Json.Formatting;
#if DEBUG
bool isOpenTestLog = true;
#else
bool isOpenTestLog = false;
#endif

// 获取当前执行文件的路径
string exePath = Assembly.GetExecutingAssembly().Location;
string exeDirectory = Path.GetDirectoryName(exePath);
// 获取上一层目录
string parentDirectory = Directory.GetParent(exeDirectory).FullName;
string reportConfigPath = $"{parentDirectory}/导出配置.xlsx";
// 设置EPPlus许可证
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string excelRelativePath = string.Empty;
string reportRelativePath = string.Empty;
using (var package = new ExcelPackage(new FileInfo(reportConfigPath)))
{
    var worksheet = package.Workbook.Worksheets[0];
    excelRelativePath  = worksheet.Cells[1, 2].Text;
    reportRelativePath = worksheet.Cells[2, 2].Text;
}
string excelsPath = Path.GetFullPath(Path.Combine(exePath, excelRelativePath));
string reportConfig = Path.GetFullPath(Path.Combine(exePath, reportRelativePath));
Log($"excelsPath: {excelsPath}", isOpenTestLog);
Log($"reportPath: {reportConfig}", isOpenTestLog);
// 获取上一层目录下的所有文件
string[] files = Directory.GetFiles(excelsPath);

var jsonData = new Dictionary<string, Dictionary<string, Dictionary<string, object>>>();
// 输出所有文件的路径和内容
foreach (string file in files)
{
    if (!file.Contains("~$"))
    {
        Log("File Path: " + file, isOpenTestLog);
        // 读取Excel文件
        var excelData = ReadExcelFile(file);
        if (excelData!=null)
        {
            jsonData.Clear();
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
            jsonData.Add(fileNameWithoutExtension, excelData);
            // 将数据转换为JSON格式
            string json = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
            // 输出JSON
            Log(json, isOpenTestLog);
            string[] reportPathAndType = reportConfig.Split("|");
            if (!Directory.Exists(reportPathAndType[0]))
                Directory.CreateDirectory(reportPathAndType[0]);
            //todo 后续可扩展添加导出数据类型
            switch (reportPathAndType[1])
            {
                case "json":
                    File.WriteAllText($"{reportPathAndType[0]}/{fileNameWithoutExtension}.json", json);
                    break;
            }
        }
    }
}
LogGreen($"转换完成！");
#if RELEASE
Console.WriteLine("按任意键关闭此窗口...");
Console.ReadKey(true);
#endif
return;

//读取excel文件数据
static Dictionary<string,Dictionary<string, object>> ReadExcelFile(string filePath)
{
    var result = new Dictionary<string, Dictionary<string, object>>();
    string fileName = Path.GetFileName(filePath);
    LogGreen(fileName);
    using (var package = new ExcelPackage(new FileInfo(filePath)))
    {
        var worksheet = package.Workbook.Worksheets[0]; // 假设数据在第一个工作表中
        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;
        int typesRowIndex = 2;
        int headersRowIndex = 3;
        int isCanWriteRowIndex = 4;
        // 读取是否写入
        var isCanWrite = new Dictionary<int,bool>();
        for (int col = 1; col <= colCount; col++)
        {
            string text = worksheet.Cells[isCanWriteRowIndex, col].Text;
            isCanWrite.Add(col,text.Equals("C"));
        }
        
        var types = new Dictionary<int, string>();
        var headers = new Dictionary<int, string>();
        foreach (var value in isCanWrite)
        {
            if (value.Value)
            {
                // 读取类型
                string typeStr = worksheet.Cells[typesRowIndex, value.Key].Text;
                if (string.IsNullOrEmpty(typeStr))
                    LogError($"数值类型行，列{value.Key},存在空值");
                types.Add(value.Key, typeStr);
                // 读取表头
                string headerStr = worksheet.Cells[headersRowIndex, value.Key].Text;
                if (string.IsNullOrEmpty(headerStr))
                    LogError($"数值名称行，列{value.Key},存在空值");
                headers.Add(value.Key, headerStr);
            }
        }
        int dataStartRowIndex = 5;
        //有效行数
        int validRowNum = 0;
        for (int row = dataStartRowIndex; row <= rowCount; row++)
        {
            string firstDataStr = worksheet.Cells[row, 1].Text;
            if (!string.IsNullOrEmpty(firstDataStr))
                validRowNum++;
        }
        if (validRowNum==0)//空表
        {
            Log($"跳过空表：{fileName}",true);
            return null;
        }
        else
        {
            var idStrList = new Dictionary<string,int>();
            for (int row = dataStartRowIndex; row <= rowCount; row++)
            {
                string idStr = worksheet.Cells[row, 1].Text;
                if (!string.IsNullOrEmpty(idStr))
                {
                    if (idStrList.ContainsKey(idStr))
                    {
                        LogError($"{fileName}表,行：{idStrList[idStr]}与行{row}id重复");
                        Environment.Exit(1);
                    }
                    var rowData = new Dictionary<string, object>();
                    foreach (var value in isCanWrite)
                    {
                        if (value.Value)
                        {
                            object oldValue = null;
                            object convertedValue = null;
                            if (worksheet.Cells[row, value.Key].Value == null)
                            {
                                oldValue = GetDefaultValue(types[value.Key]);
                                convertedValue = Convert.ChangeType(oldValue, GetValueType(types[value.Key]));
                            }
                            else
                            {
                                if (types[value.Key].Contains("[]"))//数组类型
                                {
                                    convertedValue = GetArrayTypeValue(types[value.Key], worksheet.Cells[row, value.Key].Text);
                                }
                                else
                                {
                                    oldValue = worksheet.Cells[row, value.Key].Value;
                                    convertedValue = Convert.ChangeType(oldValue, GetValueType(types[value.Key]));
                                }
                            }
                            rowData.Add(headers[value.Key], convertedValue);
                        }
                    }
                    result.Add(idStr,rowData);
                    idStrList.Add(idStr, row);
                }
            }
        }
    }
    return result;
}

//获取对应类型默认值 todo 后续可扩展相应类型默认值
static object GetDefaultValue(string type)
{
    switch (type)
    {
        case "int":
            return default(int);
        case "int[]":
            return new List<int>();
        case "long":
            return default(long);
        case "long[]":
            return new List<long>();
        case "float":
            return default(float);
        case "float[]":
            return new List<float>();
        case "double":
            return default(double);
        case "double[]":
            return new List<double>();
        case "bool":
            return default(bool);
        case "bool[]":
            return new List<bool>();
        case "char":
            return default(char);
        case "char[]":
            return new List<char>();
        case "string":
            return string.Empty;
        case "string[]":
            return new List<string>();
        default:
            return null;
    }
}
//获取对应类型 todo 后续可扩展相应类型
static Type GetValueType(string type)
{
    switch (type)
    {
        case "int":
            return typeof(int);
        case "long":
            return typeof(long);
        case "float":
            return typeof(float);
        case "double":
            return typeof(double);
        case "bool":
            return typeof(bool);
        case "char":
            return typeof(char);
        case "string":
            return typeof(string);
        case "int[]":
            return typeof(List<int>);
        case "long[]":
            return typeof(List<long>);
        case "float[]":
            return typeof(List<float>);
        case "double[]":
            return typeof(List<double>);
        case "bool[]":
            return typeof(List<bool>);
        case "char[]":
            return typeof(List<char>);
        case "string[]":
            return typeof(List<string>);
        default:
            LogError($"未添加{type}类型转换");
            return typeof(string);
    }
}
//数组类型转换 todo 后续可扩展相应类型
static object GetArrayTypeValue(string type,string value)
{
    string elementTypeStr = type.Replace("[]", "");
    Type elementType = GetValueType(elementTypeStr);
    switch (type)
    {
        case "int[]":
            return ArrayTypeSwitch<int>(elementType, value);
        case "long[]":
            return ArrayTypeSwitch<long>(elementType, value);
        case "float[]":
            return ArrayTypeSwitch<float>(elementType, value);
        case "double[]":
            return ArrayTypeSwitch<double>(elementType, value);
        case "bool[]":
            return ArrayTypeSwitch<bool>(elementType, value);
        case "char[]":
            return ArrayTypeSwitch<char>(elementType, value);
        case "string[]":
            return ArrayTypeSwitch<string>(elementType, value);
        default:
            return ArrayTypeSwitch<string>(elementType, value);
    }
}
//数组类型转换
static List<T> ArrayTypeSwitch<T>(Type elementType, string value)
{
    var valueList = new List<T>();
    var elementArray = value.Split("^");
    for (int i = 0; i < elementArray.Length; i++)
    {
        T elementValue =(T)Convert.ChangeType(elementArray[i], elementType);
        valueList.Add(elementValue);
    }
    return valueList;
}

//打印普通日志
static void Log(string msg,bool isOpen)
{
    if (isOpen)
    {
        Console.ResetColor();
        Console.WriteLine(msg);
    }
}
//打印执行正确日志
static void LogGreen(string msg)
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine(msg);
}
//打印错误日志
static void LogError(string msg)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine(msg);
}