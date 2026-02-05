using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Sparkline;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelSparklineTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
///     Note: Sparkline operations have limitations in Aspose.Cells evaluation mode,
///     so most tests use [SkippableFact] with SkipInEvaluationMode.
/// </summary>
public class ExcelSparklineToolTests : ExcelTestBase
{
    private readonly ExcelSparklineTool _tool;

    public ExcelSparklineToolTests()
    {
        _tool = new ExcelSparklineTool(SessionManager);
    }

    private string CreateWorkbookWithSparklineData(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < 5; i++) sheet.Cells[i, 0].Value = i + 1;
        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithSparkline(string fileName)
    {
        var path = CreateWorkbookWithSparklineData(fileName);
        using var workbook = new Workbook(path);
        var sheet = workbook.Worksheets[0];
        var sheetName = sheet.Name;
        var locationArea = CellArea.CreateCellArea("B1", "B1");
        sheet.SparklineGroups.Add(SparklineType.Line, $"{sheetName}!A1:A5", true, locationArea);
        workbook.Save(path);
        return path;
    }

    /// <summary>
    ///     Gets the fully-qualified data range with sheet name prefix for sparkline operations.
    /// </summary>
    private static string GetQualifiedDataRange(string path)
    {
        using var workbook = new Workbook(path);
        var sheetName = workbook.Worksheets[0].Name;
        return $"{sheetName}!A1:A5";
    }

    #region File I/O Smoke Tests

    [SkippableFact]
    public void Add_ShouldAddSparkline()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparklineData("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var dataRange = GetQualifiedDataRange(workbookPath);
        var result = _tool.Execute("add", workbookPath, dataRange: dataRange, locationRange: "B1",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("sparkline", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].SparklineGroups.Count >= 1);
    }

    [SkippableFact]
    public void Get_ShouldReturnSparklines()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparkline("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetSparklinesExcelResult>(result);
        Assert.True(data.Count >= 1);
        Assert.True(data.Items.Count >= 1);
    }

    [SkippableFact]
    public void Delete_ShouldDeleteSparkline()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparkline("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, groupIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].SparklineGroups);
    }

    [SkippableFact]
    public void SetStyle_ShouldSetSparklineStyle()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparkline("test_set_style.xlsx");
        var outputPath = CreateTestFilePath("test_set_style_output.xlsx");
        var result = _tool.Execute("set_style", workbookPath, groupIndex: 0, showHighPoint: true,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("style", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Operation Routing

    [SkippableTheory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparklineData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var dataRange = GetQualifiedDataRange(workbookPath);
        var result = _tool.Execute(operation, workbookPath, dataRange: dataRange, locationRange: "B1",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("sparkline", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithSparklineData("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Add_WithSession_ShouldAddInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparklineData("test_session_add.xlsx");
        var dataRange = GetQualifiedDataRange(workbookPath);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, dataRange: dataRange, locationRange: "B1");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("sparkline", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].SparklineGroups.Count >= 1);
    }

    [SkippableFact]
    public void Get_WithSession_ShouldGetFromMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparkline("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetSparklinesExcelResult>(result);
        Assert.True(data.Count >= 1);
        var output = GetResultOutput<GetSparklinesExcelResult>(result);
        Assert.True(output.IsSession);
    }

    [SkippableFact]
    public void Delete_WithSession_ShouldDeleteInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbookPath = CreateWorkbookWithSparkline("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, groupIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].SparklineGroups);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var pathWorkbook = CreateWorkbookWithSparklineData("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithSparkline("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var data = GetResultData<GetSparklinesExcelResult>(result);
        Assert.True(data.Count >= 1);
    }

    #endregion
}
