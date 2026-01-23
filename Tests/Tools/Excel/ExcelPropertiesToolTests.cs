using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Properties;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelPropertiesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelPropertiesToolTests : ExcelTestBase
{
    private readonly ExcelPropertiesTool _tool;

    public ExcelPropertiesToolTests()
    {
        _tool = new ExcelPropertiesTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetWorkbookProperties_ShouldReturnJsonWithAllFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_workbook.xlsx");
        var result = _tool.Execute("get_workbook_properties", workbookPath);
        var data = GetResultData<GetWorkbookPropertiesResult>(result);
        Assert.NotNull(data.Created);
        Assert.NotNull(data.Modified);
    }

    [Fact]
    public void SetWorkbookProperties_ShouldSetAllBuiltInProperties()
    {
        var workbookPath = CreateExcelWorkbook("test_set_workbook.xlsx");
        var outputPath = CreateTestFilePath("test_set_workbook_output.xlsx");
        var result = _tool.Execute("set_workbook_properties", workbookPath,
            title: "Test Title", author: "Test Author", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Workbook properties updated", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Test Title", workbook.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", workbook.BuiltInDocumentProperties.Author);
    }

    [Fact]
    public void GetSheetProperties_ShouldReturnJsonWithAllFields()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_sheet.xlsx");
        var result = _tool.Execute("get_sheet_properties", workbookPath, sheetIndex: 0);
        var data = GetResultData<GetSheetPropertiesResult>(result);
        Assert.NotNull(data.Name);
        Assert.True(data.IsVisible);
    }

    [Fact]
    public void EditSheetProperties_ShouldChangeName()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_name.xlsx");
        var outputPath = CreateTestFilePath("test_edit_name_output.xlsx");
        _tool.Execute("edit_sheet_properties", workbookPath, sheetIndex: 0, name: "NewSheetName",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("NewSheetName", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void GetSheetInfo_ShouldReturnAllSheets()
    {
        var workbookPath = CreateExcelWorkbook("test_get_info_all.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_sheet_info", workbookPath);
        var data = GetResultData<GetSheetInfoResult>(result);
        Assert.True(data.Count >= 2);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_WORKBOOK_PROPERTIES")]
    [InlineData("Get_Workbook_Properties")]
    [InlineData("get_workbook_properties")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        var data = GetResultData<GetWorkbookPropertiesResult>(result);
        Assert.NotNull(data.Created);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetWorkbookProperties_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_workbook.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.BuiltInDocumentProperties.Title = "Session Title";
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_workbook_properties", sessionId: sessionId);
        var data = GetResultData<GetWorkbookPropertiesResult>(result);
        Assert.Equal("Session Title", data.Title);
    }

    [Fact]
    public void SetWorkbookProperties_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_set_workbook.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_workbook_properties", sessionId: sessionId, title: "Updated Title");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Workbook properties updated", data.Message);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Updated Title", workbook.BuiltInDocumentProperties.Title);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_workbook_properties", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbook("test_session_file.xlsx");
        using (var wb = new Workbook(workbookPath2))
        {
            wb.BuiltInDocumentProperties.Title = "Session Title";
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get_workbook_properties", workbookPath1, sessionId);
        var data = GetResultData<GetWorkbookPropertiesResult>(result);
        Assert.Equal("Session Title", data.Title);
    }

    #endregion
}
