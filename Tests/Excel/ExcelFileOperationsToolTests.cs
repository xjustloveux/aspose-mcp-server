using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFileOperationsToolTests : ExcelTestBase
{
    private readonly ExcelFileOperationsTool _tool = new();

    #region Error Handling Tests

    [Fact]
    public async Task UnknownOperation_ShouldThrowException()
    {
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = "test.xlsx"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Create Tests

    [Fact]
    public async Task CreateWorkbook_ShouldCreateNewWorkbook()
    {
        var outputPath = CreateTestFilePath("test_create_workbook.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["path"] = outputPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count > 0);
    }

    [Fact]
    public async Task CreateWorkbook_WithSheetName_ShouldSetSheetName()
    {
        var outputPath = CreateTestFilePath("test_create_workbook_sheetname.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["path"] = outputPath,
            ["sheetName"] = "MySheet"
        };

        await _tool.ExecuteAsync(arguments);

        using var workbook = new Workbook(outputPath);
        Assert.Equal("MySheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public async Task CreateWorkbook_WithOutputPath_ShouldWork()
    {
        var outputPath = CreateTestFilePath("test_create_with_outputPath.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["outputPath"] = outputPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Convert Tests

    [Fact]
    public async Task ConvertWorkbook_ToPdf_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = "pdf"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Workbook converted to pdf format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ConvertWorkbook_ToCsv_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_csv.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.csv");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = "csv"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Workbook converted to csv format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ConvertWorkbook_ToHtml_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_html.xlsx", 2);
        var outputPath = CreateTestFilePath("test_convert_output.html");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = "html"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Workbook converted to html format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ConvertWorkbook_InvalidFormat_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_convert_invalid_output.xyz");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = "xyz"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported format", ex.Message);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xls")]
    [InlineData("ods")]
    public async Task ConvertWorkbook_AllFormats_ShouldWork(string format)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_convert_{format}.xlsx", 2);
        var outputPath = CreateTestFilePath($"test_convert_output.{format}");
        var arguments = new JsonObject
        {
            ["operation"] = "convert",
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath,
            ["format"] = format
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains($"Workbook converted to {format} format", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Merge Tests

    [Fact]
    public async Task MergeWorkbooks_ShouldMergeWorkbooks()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge2.xlsx", 2, 2);

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

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count >= 2);
    }

    [Fact]
    public async Task MergeWorkbooks_WithMergeSheets_ShouldAppendData()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge_sheets1.xlsx", 3, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_sheets2.xlsx", 3, 2);

        var outputPath = CreateTestFilePath("test_merge_sheets_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = outputPath,
            ["inputPaths"] = new JsonArray { workbook1Path, workbook2Path },
            ["mergeSheets"] = true
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Merged 2 workbooks successfully", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets);
        Assert.True(workbook.Worksheets[0].Cells.MaxDataRow >= 5);
    }

    [Fact]
    public async Task MergeWorkbooks_EmptyInputPaths_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = outputPath,
            ["inputPaths"] = new JsonArray()
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("At least one input path is required", ex.Message);
    }

    [Fact]
    public async Task MergeWorkbooks_WithOutputPath_ShouldWork()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge_op1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_op2.xlsx", 2, 2);

        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_op_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["outputPath"] = outputPath,
            ["inputPaths"] = new JsonArray { workbook1Path, workbook2Path }
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Split Tests

    [Fact]
    public async Task SplitWorkbook_ShouldSplitIntoMultipleFiles()
    {
        var workbookPath = CreateExcelWorkbook("test_split.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets[0].Cells["A1"].Value = "Data1";
            workbook.Worksheets[1].Cells["A1"].Value = "Data2";
            workbook.Worksheets[2].Cells["A1"].Value = "Data3";
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

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 3 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public async Task SplitWorkbook_WithSheetIndices_ShouldSplitOnlySpecifiedSheets()
    {
        var workbookPath = CreateExcelWorkbook("test_split_indices.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets[0].Cells["A1"].Value = "Data1";
            workbook.Worksheets[1].Cells["A1"].Value = "Data2";
            workbook.Worksheets[2].Cells["A1"].Value = "Data3";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_indices_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = workbookPath,
            ["outputDirectory"] = outputDir,
            ["sheetIndices"] = new JsonArray { 0, 2 }
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public async Task SplitWorkbook_WithDuplicateIndices_ShouldRemoveDuplicates()
    {
        var workbookPath = CreateExcelWorkbook("test_split_dup.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[0].Cells["A1"].Value = "Data1";
            workbook.Worksheets[1].Cells["A1"].Value = "Data2";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_dup_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = workbookPath,
            ["outputDirectory"] = outputDir,
            ["sheetIndices"] = new JsonArray { 0, 0, 1, 1, 0 }
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public async Task SplitWorkbook_WithCustomPattern_ShouldUsePattern()
    {
        var workbookPath = CreateExcelWorkbook("test_split_pattern.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_pattern_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = workbookPath,
            ["outputDirectory"] = outputDir,
            ["outputFileNamePattern"] = "output_{index}_{name}.xlsx"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 1 files", result);
        var files = Directory.GetFiles(outputDir, "output_*.xlsx");
        Assert.Single(files);
        Assert.Contains("output_0_", files[0]);
    }

    [Fact]
    public async Task SplitWorkbook_WithPath_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbook("test_split_path.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_path_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["path"] = workbookPath,
            ["outputDirectory"] = outputDir
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 1 files", result);
    }

    [Fact]
    public async Task SplitWorkbook_InvalidSheetIndex_ShouldSkip()
    {
        var workbookPath = CreateExcelWorkbook("test_split_invalid_idx.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_invalid_idx_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["inputPath"] = workbookPath,
            ["outputDirectory"] = outputDir,
            ["sheetIndices"] = new JsonArray { 0, 99, -1 }
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Split workbook into 1 files", result);
    }

    #endregion
}