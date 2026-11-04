using System.Security.Cryptography;
using ToolVba.Core;

if (args.Length == 5 && args[0].Equals("--bas", StringComparison.OrdinalIgnoreCase))
    return await RunBasIntegrationAsync(args[1], args[2], args[3], args[4]);

if (args.Length != 3)
{
    Console.Error.WriteLine("Cách dùng:");
    Console.Error.WriteLine("  ToolVba.IntegrationRunner <workbook.xlsx|xlsm> <vbaProject.bin> <output.xlsm>");
    Console.Error.WriteLine("  ToolVba.IntegrationRunner --bas <workbook.xlsm> <module.bas> <manifest.json> <output.xlsm>");
    return 2;
}

string inputPath;
string vbaPath;
string outputPath;
try
{
    inputPath = Path.GetFullPath(args[0]);
    vbaPath = Path.GetFullPath(args[1]);
    outputPath = Path.GetFullPath(args[2]);
}
catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
{
    Console.Error.WriteLine("FAIL: Có đường dẫn không hợp lệ hoặc quá dài.");
    return 2;
}

if (!File.Exists(inputPath) || !File.Exists(vbaPath))
{
    Console.Error.WriteLine("FAIL: Không tìm thấy workbook hoặc file VBA project.");
    return 2;
}

if (File.Exists(outputPath))
{
    Console.Error.WriteLine("FAIL: Output đã tồn tại; runner không ghi đè file.");
    return 2;
}

try
{
    var sourceHashBefore = await HashFileAsync(inputPath);
    var result = await new FileOperationService().CreateAsync(new InjectionRequest(inputPath, vbaPath, outputPath));
    var sourceHashAfter = await HashFileAsync(inputPath);

    if (!CryptographicOperations.FixedTimeEquals(sourceHashBefore, sourceHashAfter))
    {
        Console.Error.WriteLine("FAIL: SHA-256 của workbook nguồn đã thay đổi.");
        return 1;
    }

    Console.WriteLine("PASS: Workbook nguồn giữ nguyên SHA-256.");
    if (!result.Succeeded)
    {
        Console.Error.WriteLine("FAIL: " + result.Message);
        return 1;
    }

    var validation = new ExcelPackageValidator().Validate(outputPath, vbaPath);
    if (!validation.IsValid)
    {
        Console.Error.WriteLine("FAIL: " + string.Join(" ", validation.Errors));
        return 1;
    }

    Console.WriteLine("PASS: Output là package .xlsm hợp lệ, có relationship và entry VBA; SHA-256 binary khớp.");
    Console.WriteLine("Output: " + outputPath);
    foreach (var warning in result.Warnings ?? [])
        Console.WriteLine("Cảnh báo: " + warning);
    return 0;
}
catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
{
    Console.Error.WriteLine("FAIL: Không thể hoàn thành kiểm thử do lỗi truy cập file hoặc đường dẫn.");
    return 1;
}

static async Task<byte[]> HashFileAsync(string path)
{
    await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 128 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
    return await SHA256.HashDataAsync(stream);
}

static async Task<int> RunBasIntegrationAsync(string inputArg, string moduleArg, string manifestArg, string outputArg)
{
    string input;
    string module;
    string manifest;
    string output;
    try
    {
        input = Path.GetFullPath(inputArg);
        module = Path.GetFullPath(moduleArg);
        manifest = Path.GetFullPath(manifestArg);
        output = Path.GetFullPath(outputArg);
    }
    catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
    {
        Console.Error.WriteLine("FAIL: Có đường dẫn không hợp lệ hoặc quá dài.");
        return 2;
    }

    if (!File.Exists(input) || !File.Exists(module) || !File.Exists(manifest))
    {
        Console.Error.WriteLine("FAIL: Không tìm thấy workbook, module .bas hoặc manifest.");
        return 2;
    }
    if (File.Exists(output))
    {
        Console.Error.WriteLine("FAIL: Output đã tồn tại; runner không ghi đè file.");
        return 2;
    }

    try
    {
        var sourceHashBefore = await HashFileAsync(input);
        var progress = new Progress<ExcelModuleOperationProgress>(value => Console.WriteLine(value.Stage));
        var result = await new ExcelModuleInsertionService(auditLog: new SafeOperationAuditLog()).InsertAsync(
            new ExcelModuleInsertionRequest(input, module, manifest, output), progress);
        var sourceHashAfter = await HashFileAsync(input);

        if (!CryptographicOperations.FixedTimeEquals(sourceHashBefore, sourceHashAfter))
        {
            Console.Error.WriteLine("FAIL: SHA-256 của workbook nguồn đã thay đổi.");
            return 1;
        }
        Console.WriteLine("PASS: Workbook nguồn giữ nguyên SHA-256.");
        if (!result.Succeeded)
        {
            Console.Error.WriteLine("FAIL: " + result.Message);
            return 1;
        }

        Console.WriteLine("PASS: Module đã được import, component cũ được giữ và package .xlsm hợp lệ.");
        Console.WriteLine("Output: " + output);
        foreach (var warning in result.Warnings ?? [])
            Console.WriteLine("Cảnh báo: " + warning);
        return 0;
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
    {
        Console.Error.WriteLine("FAIL: Không thể hoàn thành kiểm thử do lỗi truy cập file hoặc đường dẫn.");
        return 1;
    }
}
