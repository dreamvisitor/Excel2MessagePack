// 确保在项目中安装了 EPPlus 包
// dotnet add package EPPlus

using Excel2MessagePack;
using OfficeOpenXml;


string iniFilePath = "setting.ini";

Console.WriteLine($"当前工作目录：{Directory.GetCurrentDirectory()}");
Console.WriteLine($"配置文件路径：{Path.GetFullPath(iniFilePath)}");
Console.WriteLine("=====================================================");

// Step 1: 读取 setting.ini
Setting? settings = ReadIni.ReadSettings(iniFilePath);
if (settings == null)
{
    Console.WriteLine($"读取setting.ini失败，已创建在 {Path.GetFullPath(iniFilePath)}，请修改配置后重新运行。");
    return;
}

// 设置 EPPlus 的 LicenseContext
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

if (!Directory.Exists(settings.TargetFolder))
{
    Console.WriteLine($"目标文件夹[{settings.TargetFolder}]不存在，请检查配置。");
    return;
}

if (!Directory.Exists(settings.OutputFolder))
{
    Console.WriteLine($"输出文件夹[{settings.OutputFolder}]不存在，请检查配置。");
    return;
}

if (!Directory.Exists(settings.SourceCodeFolder))
{
    Console.WriteLine($"源码文件夹[{settings.SourceCodeFolder}]不存在，请检查配置。");
    return;
}

// Step 3: 读取 Excel 文件
var excelFiles = Directory.GetFiles(settings.TargetFolder, "*.xlsx");

// Step 4: 处理每个Excel文件
foreach (var excelFile in excelFiles)
{
    Console.WriteLine($"********* 处理文件：{excelFile}");

    using var package = new ExcelPackage(new FileInfo(excelFile));
    foreach (var worksheet in package.Workbook.Worksheets)
    {
        // Step 5: 处理每个工作表
        ProcessExcel processExcel = new(worksheet, settings);
        processExcel.Process();
    }
}

