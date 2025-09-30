// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;

Console.WriteLine("Welcome to filepath generator...");
// Set the directory to search
string searchDirectory = @"C:\REPOS\MC_8_DEV\Acceptance\FeLinesTests";
string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FilesNames.xlsx");
string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FoundFilePaths.txt");
Console.WriteLine("searchDirectory : " + searchDirectory);
Console.WriteLine("excelPath : " + excelPath);
Console.WriteLine("outputPath : " + outputPath);

if (!File.Exists(excelPath))
{
    Console.WriteLine($"Excel file not found: {excelPath}");
    return;
}

// 1. Read file names from Excel
var fileNames = new List<string>();
using (var workbook = new XLWorkbook(excelPath))
{
    var ws = workbook.Worksheet(1);
    foreach (var row in ws.RowsUsed())
    {
        var name = row.Cell(1).GetString().Trim();
        if (!string.IsNullOrEmpty(name))
            fileNames.Add(name);
    }
}

Console.WriteLine("Press Enter to start searching for files...");
Console.ReadLine();

// 2. Search the folder and subfolders, displaying every checked file
var foundFiles = new List<string>();
var allFiles = Directory.GetFiles(searchDirectory, "*", SearchOption.AllDirectories);

foreach (var fileName in fileNames)
{
    Console.WriteLine($"Checking file: {fileName}");
    foreach (var file in allFiles)
    {
        if (string.Equals(Path.GetFileName(file), fileName, StringComparison.OrdinalIgnoreCase))
        {
            foundFiles.Add($"\"{file}\"");
        }
    }
}

// 3. Save to text file
File.WriteAllLines(outputPath, foundFiles);

Console.WriteLine($"Found files have been saved to: {outputPath}");
Console.WriteLine($"Press any key to end program..");
Console.ReadLine();