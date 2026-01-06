using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelFileOperationsToolTests : ExcelTestBase
{
    private readonly ExcelFileOperationsTool _tool;

    public ExcelFileOperationsToolTests()
    {
        _tool = new ExcelFileOperationsTool(SessionManager);
    }

    #region General

    [Fact]
    public void Create_ShouldCreateNewWorkbook()
    {
        var outputPath = CreateTestFilePath("test_create.xlsx");
        var result = _tool.Execute("create", path: outputPath);
        Assert.StartsWith("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count > 0);
    }

    [Fact]
    public void Create_WithSheetName_ShouldSetSheetName()
    {
        var outputPath = CreateTestFilePath("test_create_sheet.xlsx");
        _tool.Execute("create", path: outputPath, sheetName: "MySheet");
        using var workbook = new Workbook(outputPath);
        Assert.Equal("MySheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void Create_WithOutputPath_ShouldWork()
    {
        var outputPath = CreateTestFilePath("test_create_output.xlsx");
        var result = _tool.Execute("create", outputPath: outputPath);
        Assert.StartsWith("Excel workbook created successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_ToPdf_ShouldConvert()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_pdf.xlsx", 3);
        var outputPath = CreateTestFilePath("test_convert_output.pdf");
        var result = _tool.Execute("convert", inputPath: workbookPath, outputPath: outputPath, format: "pdf");
        Assert.StartsWith("Workbook from", result);
        Assert.Contains("converted to pdf format", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("html")]
    [InlineData("xlsx")]
    [InlineData("xls")]
    [InlineData("ods")]
    public void Convert_AllFormats_ShouldWork(string format)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_convert_{format}.xlsx", 2);
        var outputPath = CreateTestFilePath($"test_convert_output.{format}");
        var result = _tool.Execute("convert", inputPath: workbookPath, outputPath: outputPath, format: format);
        Assert.StartsWith("Workbook from", result);
        Assert.Contains($"converted to {format} format", result);
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
        Assert.StartsWith("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets.Count >= 2);
    }

    [SkippableFact]
    public void Merge_WithMergeSheets_ShouldAppendData()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Merging sheets exceeds row limit in evaluation mode");
        var workbook1Path = CreateExcelWorkbookWithData("test_merge_sheets1.xlsx", 3, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_sheets2.xlsx", 3, 2);
        var outputPath = CreateTestFilePath("test_merge_sheets_output.xlsx");
        var result = _tool.Execute("merge", path: outputPath, inputPaths: [workbook1Path, workbook2Path],
            mergeSheets: true);
        Assert.StartsWith("Merged 2 workbooks successfully", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets);
        Assert.True(workbook.Worksheets[0].Cells.MaxDataRow >= 5);
    }

    [Fact]
    public void Merge_WithOutputPath_ShouldWork()
    {
        var workbook1Path = CreateExcelWorkbookWithData("test_merge_op1.xlsx", 2, 2);
        var workbook2Path = CreateExcelWorkbookWithData("test_merge_op2.xlsx", 2, 2);
        using (var wb2 = new Workbook(workbook2Path))
        {
            wb2.Worksheets[0].Name = "Sheet2";
            wb2.Save(workbook2Path);
        }

        var outputPath = CreateTestFilePath("test_merge_op_output.xlsx");
        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [workbook1Path, workbook2Path]);
        Assert.StartsWith("Merged 2 workbooks successfully", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Split_ShouldSplitIntoMultipleFiles()
    {
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
        var result = _tool.Execute("split", inputPath: workbookPath, outputDirectory: outputDir);
        Assert.StartsWith("Split workbook into 3 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(3, files.Length);
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
        Assert.StartsWith("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public void Split_WithDuplicateIndices_ShouldRemoveDuplicates()
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
        var result = _tool.Execute("split", inputPath: workbookPath, outputDirectory: outputDir,
            sheetIndices: [0, 0, 1, 1, 0]);
        Assert.StartsWith("Split workbook into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.xlsx");
        Assert.Equal(2, files.Length);
    }

    [SkippableFact]
    public void Split_WithCustomPattern_ShouldUsePattern()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Split operation may add evaluation sheet");
        var workbookPath = CreateExcelWorkbook("test_split_pattern.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_pattern_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: workbookPath, outputDirectory: outputDir,
            outputFileNamePattern: "output_{index}_{name}.xlsx");
        Assert.StartsWith("Split workbook into 1 files", result);
        var files = Directory.GetFiles(outputDir, "output_*.xlsx");
        Assert.Single(files);
        Assert.Contains("output_0_", files[0]); // Verify pattern was used
    }

    [SkippableFact]
    public void Split_WithPath_ShouldWork()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Split operation may add evaluation sheet");
        var workbookPath = CreateExcelWorkbook("test_split_path.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_path_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", path: workbookPath, outputDirectory: outputDir);
        Assert.StartsWith("Split workbook into 1 files", result);
    }

    [Fact]
    public void Split_InvalidSheetIndex_ShouldSkip()
    {
        var workbookPath = CreateExcelWorkbook("test_split_invalid_idx.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Data";
            workbook.Save(workbookPath);
        }

        var outputDir = Path.Combine(TestDir, "split_invalid_idx_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("split", inputPath: workbookPath, outputDirectory: outputDir,
            sheetIndices: [0, 99, -1]);
        Assert.StartsWith("Split workbook into 1 files", result);
    }

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive_Create(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.xlsx");
        var result = _tool.Execute(operation, path: outputPath);
        Assert.StartsWith("Excel workbook created successfully", result);
    }

    [Theory]
    [InlineData("CONVERT")]
    [InlineData("Convert")]
    [InlineData("convert")]
    public void Operation_ShouldBeCaseInsensitive_Convert(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_convert_{operation}.xlsx", 2);
        var outputPath = CreateTestFilePath($"test_case_convert_{operation}_output.csv");
        var result = _tool.Execute(operation, inputPath: workbookPath, outputPath: outputPath, format: "csv");
        Assert.StartsWith("Workbook from", result);
        Assert.Contains("converted to csv format", result);
    }

    [Theory]
    [InlineData("MERGE")]
    [InlineData("Merge")]
    [InlineData("merge")]
    public void Operation_ShouldBeCaseInsensitive_Merge(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_merge_{operation}.xlsx", 2);
        var outputPath = CreateTestFilePath($"test_case_merge_{operation}_output.xlsx");
        var result = _tool.Execute(operation, path: outputPath, inputPaths: [workbookPath]);
        Assert.StartsWith("Merged 1 workbooks successfully", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", "test.xlsx"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Create_WithNoPath_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("create"));
        Assert.Contains("path or outputPath is required", ex.Message);
    }

    [Fact]
    public void Convert_WithNoInputPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_convert_no_input.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", outputPath: outputPath, format: "pdf"));
        Assert.Contains("inputPath or sessionId is required", ex.Message);
    }

    [Fact]
    public void Convert_WithNoOutputPath_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_no_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: workbookPath, format: "pdf"));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void Convert_WithNoFormat_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_no_format.xlsx");
        var outputPath = CreateTestFilePath("test_convert_no_format.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: workbookPath, outputPath: outputPath));
        Assert.Contains("format is required", ex.Message);
    }

    [Fact]
    public void Convert_WithInvalidFormat_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_convert_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_convert_invalid_output.xyz");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("convert", inputPath: workbookPath, outputPath: outputPath, format: "xyz"));
        Assert.Contains("Unsupported format", ex.Message);
    }

    [Fact]
    public void Merge_WithNoPath_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_no_path.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", inputPaths: [workbookPath]));
        Assert.Contains("path or outputPath is required", ex.Message);
    }

    [Fact]
    public void Merge_WithEmptyInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", path: outputPath, inputPaths: Array.Empty<string>()));
        Assert.Contains("At least one input path is required", ex.Message);
    }

    [Fact]
    public void Merge_WithNullInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_null.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", path: outputPath, inputPaths: null));
        Assert.Contains("At least one input path is required", ex.Message);
    }

    [Fact]
    public void Split_WithNoInputPath_ShouldThrowArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "split_no_input");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", outputDirectory: outputDir));
        Assert.Contains("inputPath, path, or sessionId is required", ex.Message);
    }

    [Fact]
    public void Split_WithNoOutputDirectory_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_split_no_dir.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", inputPath: workbookPath));
        Assert.Contains("outputDirectory is required", ex.Message);
    }

    #endregion

    #region Session

    [SkippableFact]
    public void Convert_WithSessionId_ShouldConvertFromMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Session convert may have evaluation limitations");
        var workbookPath = CreateExcelWorkbookWithData("test_session_convert.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var outputPath = CreateTestFilePath("test_session_convert_output.csv");
        var result = _tool.Execute("convert", sessionId, outputPath: outputPath, format: "csv");
        Assert.StartsWith("Workbook from", result);
        Assert.Contains("converted to csv format", result);
        Assert.Contains($"session {sessionId}", result); // Verify session was used
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
        Assert.StartsWith("Split workbook into", result);
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

    [Fact]
    public void Split_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputDir = Path.Combine(TestDir, "split_invalid_session");
        Directory.CreateDirectory(outputDir);
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("split", "invalid_session", outputDirectory: outputDir));
    }

    #endregion
}