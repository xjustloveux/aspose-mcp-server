using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFileOperationsToolTests : ExcelTestBase
{
    private readonly ExcelFileOperationsTool _tool = new();

    [Fact]
    public async Task CreateWorkbook_ShouldCreateNewWorkbook()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_workbook.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["path"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Workbook should be created");
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count > 0, "Workbook should have at least one worksheet");
    }

    [Fact]
    public async Task CreateWorkbook_WithSheetName_ShouldSetSheetName()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_workbook_sheetname.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["path"] = outputPath,
            ["sheetName"] = "MySheet"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        Assert.Equal("MySheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public async Task ConvertWorkbook_ShouldConvertToPdf()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_convert.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = "pdf"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task MergeWorkbooks_ShouldMergeWorkbooks()
    {
        // Arrange
        var workbook1Path = CreateExcelWorkbookWithData("test_merge1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge2.xlsx", 2, 2);
        // Rename sheets to avoid conflicts
        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = outputPath,
            ["inputPaths"] = new JsonArray { workbook1Path, workbook2Path }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Merged workbook should be created");
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count >= 2, "Merged workbook should have multiple sheets");
    }

    [Fact]
    public async Task SplitWorkbook_ShouldSplitIntoMultipleFiles()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_split.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            // Add some data to avoid zero row/column error
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Worksheets[1].Cells["A1"].Value = "Data";
            workbook.Worksheets[2].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = workbookPath,
            ["outputDirectory"] = outputDir
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length >= 2, "Should create multiple files for split sheets");
    }
}