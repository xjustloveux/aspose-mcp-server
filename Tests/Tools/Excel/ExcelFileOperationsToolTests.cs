using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelFileOperationsToolTests : ExcelTestBase
{
    private readonly ExcelFileOperationsTool _tool = new();

    #region General Tests

    [Fact]
    public void CreateWorkbook_ShouldCreateNewWorkbook()
    {
        var outputPath = CreateTestFilePath("test_create_workbook.xlsx");

        var result = _tool.Execute(
            "create",
            outputPath);

        Assert.Contains("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count > 0);
    }

    [Fact]
    public void CreateWorkbook_WithSheetName_ShouldSetSheetName()
    {
        var outputPath = CreateTestFilePath("test_create_workbook_sheetname.xlsx");

        _tool.Execute(
            "create",
            outputPath,
            sheetName: "MySheet");

        using var workbook = new Workbook(outputPath);
        Assert.Equal("MySheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void CreateWorkbook_WithOutputPath_ShouldWork()
    {
        var outputPath = CreateTestFilePath("test_create_with_outputPath.xlsx");

        var result = _tool.Execute(
            "create",
            outputPath: outputPath);

        Assert.Contains("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ConvertWorkbook_ToPdf_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.pdf");

        var result = _tool.Execute(
            "convert",
            inputPath: workbookPath,
            outputPath: outputPath,
            format: "pdf");

        Assert.Contains("Workbook converted to pdf format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ConvertWorkbook_ToCsv_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_csv.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.csv");

        var result = _tool.Execute(
            "convert",
            inputPath: workbookPath,
            outputPath: outputPath,
            format: "csv");

        Assert.Contains("Workbook converted to csv format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ConvertWorkbook_ToHtml_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_html.xlsx", 2);
        var outputPath = CreateTestFilePath("test_convert_output.html");

        var result = _tool.Execute(
            "convert",
            inputPath: workbookPath,
            outputPath: outputPath,
            format: "html");

        Assert.Contains("Workbook converted to html format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xls")]
    [InlineData("ods")]
    public void ConvertWorkbook_AllFormats_ShouldWork(string format)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_convert_{format}.xlsx", 2);
        var outputPath = CreateTestFilePath($"test_convert_output.{format}");

        var result = _tool.Execute(
            "convert",
            inputPath: workbookPath,
            outputPath: outputPath,
            format: format);

        Assert.Contains($"Workbook converted to {format} format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void MergeWorkbooks_ShouldMergeWorkbooks()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge2.xlsx", 2, 2);

        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_output.xlsx");

        var result = _tool.Execute(
            "merge",
            outputPath,
            inputPaths: [workbook1Path, workbook2Path]);

        Assert.Contains("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count >= 2);
    }

    [SkippableFact]
    public void MergeWorkbooks_WithMergeSheets_ShouldAppendData()
    {
        // Skip in evaluation mode - merging sheets exceeds row limit in evaluation mode
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Merging sheets exceeds row limit in evaluation mode");

        var workbook1Path = CreateExcelWorkbookWithData("test_merge_sheets1.xlsx", 3, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_sheets2.xlsx", 3, 2);

        var outputPath = CreateTestFilePath("test_merge_sheets_output.xlsx");

        var result = _tool.Execute(
            "merge",
            outputPath,
            inputPaths: [workbook1Path, workbook2Path],
            mergeSheets: true);

        Assert.Contains("Merged 2 workbooks successfully", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets);
        Assert.True(workbook.Worksheets[0].Cells.MaxDataRow >= 5);
    }

    [Fact]
    public void MergeWorkbooks_WithOutputPath_ShouldWork()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge_op1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_op2.xlsx", 2, 2);

        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_op_output.xlsx");

        var result = _tool.Execute(
            "merge",
            outputPath: outputPath,
            inputPaths: [workbook1Path, workbook2Path]);

        Assert.Contains("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void SplitWorkbook_ShouldSplitIntoMultipleFiles()
    {
        // Skip in evaluation mode - 3 sheets exceeds evaluation limit
        SkipInEvaluationMode(AsposeLibraryType.Cells, "3 sheets exceeds evaluation limit");

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

        var result = _tool.Execute(
            "split",
            inputPath: workbookPath,
            outputDirectory: outputDir);

        Assert.Contains("Split workbook into 3 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public void SplitWorkbook_WithSheetIndices_ShouldSplitOnlySpecifiedSheets()
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

        var result = _tool.Execute(
            "split",
            inputPath: workbookPath,
            outputDirectory: outputDir,
            sheetIndices: [0, 2]);

        Assert.Contains("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public void SplitWorkbook_WithDuplicateIndices_ShouldRemoveDuplicates()
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

        var result = _tool.Execute(
            "split",
            inputPath: workbookPath,
            outputDirectory: outputDir,
            sheetIndices: [0, 0, 1, 1, 0]);

        Assert.Contains("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [SkippableFact]
    public void SplitWorkbook_WithCustomPattern_ShouldUsePattern()
    {
        // Skip in evaluation mode - split operation may add evaluation sheet
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Split operation may add evaluation sheet");

        var workbookPath = CreateExcelWorkbook("test_split_pattern.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_pattern_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute(
            "split",
            inputPath: workbookPath,
            outputDirectory: outputDir,
            outputFileNamePattern: "output_{index}_{name}.xlsx");

        Assert.Contains("Split workbook into 1 files", result);
        var files = Directory.GetFiles(outputDir, "output_*.xlsx");
        Assert.Single(files);
        Assert.Contains("output_0_", files[0]);
    }

    [SkippableFact]
    public void SplitWorkbook_WithPath_ShouldWork()
    {
        // Skip in evaluation mode - split operation may add evaluation sheet
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Split operation may add evaluation sheet");

        var workbookPath = CreateExcelWorkbook("test_split_path.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_path_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute(
            "split",
            workbookPath,
            outputDirectory: outputDir);

        Assert.Contains("Split workbook into 1 files", result);
    }

    [Fact]
    public void SplitWorkbook_InvalidSheetIndex_ShouldSkip()
    {
        var workbookPath = CreateExcelWorkbook("test_split_invalid_idx.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_invalid_idx_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute(
            "split",
            inputPath: workbookPath,
            outputDirectory: outputDir,
            sheetIndices: [0, 99, -1]);

        Assert.Contains("Split workbook into 1 files", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void UnknownOperation_ShouldThrowException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unknown",
            "test.xlsx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void ConvertWorkbook_InvalidFormat_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_convert_invalid_output.xyz");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "convert",
            inputPath: workbookPath,
            outputPath: outputPath,
            format: "xyz"));
        Assert.Contains("Unsupported format", ex.Message);
    }

    [Fact]
    public void MergeWorkbooks_EmptyInputPaths_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            outputPath,
            inputPaths: Array.Empty<string>()));
        Assert.Contains("At least one input path is required", ex.Message);
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}