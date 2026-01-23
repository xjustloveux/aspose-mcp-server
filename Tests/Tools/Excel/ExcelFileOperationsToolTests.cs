using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelFileOperationsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelFileOperationsToolTests : ExcelTestBase
{
    private readonly ExcelFileOperationsTool _tool;

    public ExcelFileOperationsToolTests()
    {
        _tool = new ExcelFileOperationsTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Create_ShouldCreateNewWorkbook()
    {
        var outputPath = CreateTestFilePath("test_create.xlsx");
        var result = _tool.Execute("create", path: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Excel workbook created successfully", data.Message);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count > 0);
    }

    [Fact]
    public void Convert_ToPdf_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_pdf.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var result = _tool.Execute("convert", inputPath: workbookPath, outputPath: outputPath, format: "pdf");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Workbook from", data.Message);
        Assert.Contains("converted to pdf format", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Merge_ShouldMergeWorkbooks()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge2.xlsx", 2, 2);
        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_output.xlsx");
        var result = _tool.Execute("merge", path: outputPath, inputPaths: [workbook1Path, workbook2Path]);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Merged 2 workbooks successfully", data.Message);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count >= 2);
    }

    [Fact]
    public void Split_WithSheetIndices_ShouldSplitOnlySpecifiedSheets()
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
        var result = _tool.Execute("split", inputPath: workbookPath, outputDirectory: outputDir, sheetIndices: [0, 2]);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Split workbook into 2 files", data.Message);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.xlsx");
        var result = _tool.Execute(operation, path: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Excel workbook created successfully", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", "test.xlsx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Convert_WithSessionId_ShouldConvertFromMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Session convert may have evaluation limitations");
        var workbookPath = CreateExcelWorkbookWithData("test_session_convert.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var outputPath = CreateTestFilePath("test_session_convert_output.csv");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "csv");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Workbook from", data.Message);
        Assert.Contains("converted to csv format", data.Message);
        Assert.Contains($"session {sessionId}", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Split_WithSessionId_ShouldSplitFromMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Session split may have evaluation limitations");
        var workbookPath = CreateExcelWorkbook("test_session_split.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var outputDir = Path.Combine(TestDir, "split_session_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", sessionId, outputDirectory: outputDir);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Split workbook into", data.Message);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.True(files.Length >= 1);
    }

    [Fact]
    public void Convert_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session.csv");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("convert", "invalid_session", outputPath: outputPath, format: "csv"));
    }

    #endregion
}
