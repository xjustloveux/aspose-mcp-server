using Aspose.Cells;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Excel.DataImportExport;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelDataImportExportTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelDataImportExportToolTests : ExcelTestBase
{
    private readonly ExcelDataImportExportTool _tool;

    public ExcelDataImportExportToolTests()
    {
        _tool = new ExcelDataImportExportTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void ImportJson_ShouldImportData()
    {
        var workbookPath = CreateExcelWorkbook("test_import_json.xlsx");
        var outputPath = CreateTestFilePath("test_import_json_output.xlsx");
        var jsonData = "[{\"name\":\"John\",\"age\":30},{\"name\":\"Jane\",\"age\":25}]";
        var result = _tool.Execute("import_json", workbookPath, jsonData: jsonData, outputPath: outputPath);
        var data = GetResultData<ImportExcelResult>(result);
        Assert.True(data.RowCount >= 2);
        Assert.True(data.ColumnCount >= 2);
        Assert.Contains("import", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        var sheet = workbook.Worksheets[0];
        Assert.NotNull(sheet.Cells["A1"].Value);
    }

    [Fact]
    public void ImportArray_ShouldImportData()
    {
        var workbookPath = CreateExcelWorkbook("test_import_array.xlsx");
        var outputPath = CreateTestFilePath("test_import_array_output.xlsx");
        var arrayData = "A,B,C;1,2,3;4,5,6";
        var result = _tool.Execute("import_array", workbookPath, arrayData: arrayData, outputPath: outputPath);
        var data = GetResultData<ImportExcelResult>(result);
        Assert.True(data.RowCount >= 3);
        Assert.Equal(3, data.ColumnCount);
        Assert.Contains("import", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        var sheet = workbook.Worksheets[0];
        Assert.Equal("A", sheet.Cells["A1"].StringValue);
        Assert.Equal("B", sheet.Cells["B1"].StringValue);
    }

    [Fact]
    public void ExportCsv_ShouldExportFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_export_csv.xlsx");
        var outputPath = CreateTestFilePath("test_export_csv_output.csv");
        var result = _tool.Execute("export_csv", workbookPath, outputPath: outputPath);
        var data = GetResultData<ExportExcelResult>(result);
        Assert.Contains("export", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.True(File.Exists(outputPath));
        var csvContent = File.ReadAllText(outputPath);
        Assert.NotEmpty(csvContent);
    }

    [Fact]
    public void ExportRangeImage_ShouldExportImage()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_export_image.xlsx");
        var outputPath = CreateTestFilePath("test_export_image_output.png");
        var result = _tool.Execute("export_range_image", workbookPath, outputPath: outputPath);
        var data = GetResultData<ExportExcelResult>(result);
        Assert.Contains("export", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("IMPORT_JSON")]
    [InlineData("Import_Json")]
    [InlineData("import_json")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var jsonData = "[{\"x\":1}]";
        var result = _tool.Execute(operation, workbookPath, jsonData: jsonData, outputPath: outputPath);
        var data = GetResultData<ImportExcelResult>(result);
        Assert.Contains("import", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("import_json", jsonData: "[{\"x\":1}]"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void ImportJson_WithSession_ShouldImportInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_import_json.xlsx");
        var sessionId = OpenSession(workbookPath);
        var jsonData = "[{\"name\":\"SessionTest\",\"value\":42}]";
        var result = _tool.Execute("import_json", sessionId: sessionId, jsonData: jsonData);
        var data = GetResultData<ImportExcelResult>(result);
        Assert.Contains("import", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<ImportExcelResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void ImportArray_WithSession_ShouldImportInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_import_array.xlsx");
        var sessionId = OpenSession(workbookPath);
        var arrayData = "X,Y;10,20;30,40";
        var result = _tool.Execute("import_array", sessionId: sessionId, arrayData: arrayData);
        var data = GetResultData<ImportExcelResult>(result);
        Assert.Contains("import", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<ImportExcelResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].StringValue);
    }

    [Fact]
    public void ExportCsv_WithSession_ShouldExportFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_export_csv.xlsx");
        var sessionId = OpenSession(workbookPath);
        var outputPath = CreateTestFilePath("test_session_export_csv_output.csv");
        var result = _tool.Execute("export_csv", sessionId: sessionId, outputPath: outputPath);
        var data = GetResultData<ExportExcelResult>(result);
        Assert.Contains("export", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("import_json", sessionId: "invalid_session", jsonData: "[{\"x\":1}]"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateExcelWorkbookWithData("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var outputPath = CreateTestFilePath("test_prefer_session_output.csv");
        var result = _tool.Execute("export_csv", pathWorkbook, sessionId, outputPath);
        Assert.IsType<FinalizedResult<ExportExcelResult>>(result);
        Assert.True(File.Exists(outputPath));
        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("R1C1", csvContent);
    }

    #endregion
}
