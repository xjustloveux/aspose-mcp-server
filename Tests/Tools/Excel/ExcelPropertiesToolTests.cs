using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelPropertiesToolTests : ExcelTestBase
{
    private readonly ExcelPropertiesTool _tool;

    public ExcelPropertiesToolTests()
    {
        _tool = new ExcelPropertiesTool(SessionManager);
    }

    #region General

    [Fact]
    public void GetWorkbookProperties_ShouldReturnJsonWithAllFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_workbook.xlsx");
        var result = _tool.Execute("get_workbook_properties", workbookPath);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.True(root.TryGetProperty("title", out _));
        Assert.True(root.TryGetProperty("author", out _));
        Assert.True(root.TryGetProperty("totalSheets", out _));
        Assert.True(root.TryGetProperty("customProperties", out _));
    }

    [Fact]
    public void GetWorkbookProperties_WithSetProperties_ShouldReturnCorrectValues()
    {
        var workbookPath = CreateExcelWorkbook("test_get_workbook_values.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.BuiltInDocumentProperties.Title = "Test Title";
            workbook.BuiltInDocumentProperties.Author = "Test Author";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_workbook_properties", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal("Test Title", json.RootElement.GetProperty("title").GetString());
        Assert.Equal("Test Author", json.RootElement.GetProperty("author").GetString());
    }

    [Fact]
    public void SetWorkbookProperties_ShouldSetAllBuiltInProperties()
    {
        var workbookPath = CreateExcelWorkbook("test_set_workbook.xlsx");
        var outputPath = CreateTestFilePath("test_set_workbook_output.xlsx");
        var result = _tool.Execute("set_workbook_properties", workbookPath,
            title: "Test Title", subject: "Test Subject", author: "Test Author",
            keywords: "test,keywords", comments: "Test Comments", category: "Test Category",
            company: "Test Company", manager: "Test Manager", outputPath: outputPath);
        Assert.Contains("Workbook properties updated", result);
        using var workbook = new Workbook(outputPath);
        var props = workbook.BuiltInDocumentProperties;
        Assert.Equal("Test Title", props.Title);
        Assert.Equal("Test Subject", props.Subject);
        Assert.Equal("Test Author", props.Author);
        Assert.Equal("test,keywords", props.Keywords);
        Assert.Equal("Test Comments", props.Comments);
        Assert.Equal("Test Category", props.Category);
        Assert.Equal("Test Company", props.Company);
        Assert.Equal("Test Manager", props.Manager);
    }

    [Fact]
    public void SetWorkbookProperties_WithCustomProperties_ShouldAddNewProperties()
    {
        var workbookPath = CreateExcelWorkbook("test_set_custom.xlsx");
        var outputPath = CreateTestFilePath("test_set_custom_output.xlsx");
        var customPropsJson = new JsonObject
        {
            ["CustomProp1"] = "Value1",
            ["CustomProp2"] = "Value2"
        }.ToJsonString();
        _tool.Execute("set_workbook_properties", workbookPath, customProperties: customPropsJson,
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var customProps = workbook.CustomDocumentProperties;
        Assert.Equal("Value1", customProps["CustomProp1"].Value?.ToString());
        Assert.Equal("Value2", customProps["CustomProp2"].Value?.ToString());
    }

    [Fact]
    public void SetWorkbookProperties_WithExistingCustomProperty_ShouldUpdateValue()
    {
        var workbookPath = CreateExcelWorkbook("test_update_custom.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.CustomDocumentProperties.Add("ExistingProp", "OldValue");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_update_custom_output.xlsx");
        var customPropsJson = new JsonObject { ["ExistingProp"] = "NewValue" }.ToJsonString();
        _tool.Execute("set_workbook_properties", workbookPath, customProperties: customPropsJson,
            outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("NewValue", resultWorkbook.CustomDocumentProperties["ExistingProp"].Value?.ToString());
    }

    [Fact]
    public void GetSheetProperties_ShouldReturnJsonWithAllFields()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_sheet.xlsx");
        var result = _tool.Execute("get_sheet_properties", workbookPath, sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.True(root.TryGetProperty("name", out _));
        Assert.True(root.TryGetProperty("index", out _));
        Assert.True(root.TryGetProperty("isVisible", out _));
        Assert.True(root.TryGetProperty("dataRowCount", out _));
        Assert.True(root.TryGetProperty("dataColumnCount", out _));
        Assert.True(root.TryGetProperty("printSettings", out _));
        Assert.Equal(5, root.GetProperty("dataRowCount").GetInt32());
        Assert.Equal(3, root.GetProperty("dataColumnCount").GetInt32());
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
    public void EditSheetProperties_ShouldChangeVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_visibility.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_visibility_output.xlsx");
        _tool.Execute("edit_sheet_properties", workbookPath, sheetIndex: 0, isVisible: false, outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].IsVisible);
    }

    [Fact]
    public void EditSheetProperties_ShouldChangeTabColor()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_color.xlsx");
        var outputPath = CreateTestFilePath("test_edit_color_output.xlsx");
        _tool.Execute("edit_sheet_properties", workbookPath, sheetIndex: 0, tabColor: "#FF0000",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var tabColor = workbook.Worksheets[0].TabColor;
        Assert.Equal(255, tabColor.R);
        Assert.Equal(0, tabColor.G);
        Assert.Equal(0, tabColor.B);
    }

    [Fact]
    public void GetSheetInfo_ShouldReturnAllSheets()
    {
        var workbookPath = CreateExcelWorkbook("test_get_info_all.xlsx");
        int expectedSheetCount;
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Save(workbookPath);
            expectedSheetCount = workbook.Worksheets.Count;
        }

        var result = _tool.Execute("get_sheet_info", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(expectedSheetCount, root.GetProperty("count").GetInt32());
        Assert.Equal(expectedSheetCount, root.GetProperty("totalWorksheets").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(expectedSheetCount, items.GetArrayLength());
    }

    [Fact]
    public void GetSheetInfo_WithSheetIndex_ShouldReturnSingleSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_get_info_single.xlsx");
        int totalSheetCount;
        int sheet2Index;
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Save(workbookPath);
            totalSheetCount = workbook.Worksheets.Count;
            sheet2Index = workbook.Worksheets["Sheet2"].Index;
        }

        var result = _tool.Execute("get_sheet_info", workbookPath, targetSheetIndex: sheet2Index);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.Equal(totalSheetCount, root.GetProperty("totalWorksheets").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(1, items.GetArrayLength());
        Assert.Equal("Sheet2", items[0].GetProperty("name").GetString());
    }

    [Fact]
    public void GetSheetInfo_ShouldReturnCorrectDataCounts()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_info_data.xlsx", 10, 5);
        var result = _tool.Execute("get_sheet_info", workbookPath, sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        var sheet = json.RootElement.GetProperty("items")[0];
        Assert.Equal(10, sheet.GetProperty("dataRowCount").GetInt32());
        Assert.Equal(5, sheet.GetProperty("dataColumnCount").GetInt32());
    }

    [Theory]
    [InlineData("GET_WORKBOOK_PROPERTIES")]
    [InlineData("Get_Workbook_Properties")]
    [InlineData("get_workbook_properties")]
    public void Operation_ShouldBeCaseInsensitive_GetWorkbookProperties(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("title", result);
    }

    [Theory]
    [InlineData("GET_SHEET_INFO")]
    [InlineData("Get_Sheet_Info")]
    [InlineData("get_sheet_info")]
    public void Operation_ShouldBeCaseInsensitive_GetSheetInfo(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("items", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void GetSheetProperties_WithInvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_sheet_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_sheet_properties", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetSheetProperties_WithNegativeIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_sheet_negative.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_sheet_properties", workbookPath, sheetIndex: -1));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditSheetProperties_WithInvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_invalid.xlsx");
        var outputPath = CreateTestFilePath("test_edit_invalid_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_sheet_properties", workbookPath, sheetIndex: 99, name: "NewName",
                outputPath: outputPath));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetSheetInfo_WithInvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_info_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_sheet_info", workbookPath, targetSheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get_workbook_properties", ""));
    }

    [Fact]
    public void Execute_WithEmptyOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_empty_op.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("", workbookPath));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_workbook_properties"));
    }

    #endregion

    #region Session

    [Fact]
    public void GetWorkbookProperties_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_workbook.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.BuiltInDocumentProperties.Title = "Session Title";
            workbook.BuiltInDocumentProperties.Author = "Session Author";
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_workbook_properties", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal("Session Title", json.RootElement.GetProperty("title").GetString());
        Assert.Equal("Session Author", json.RootElement.GetProperty("author").GetString());
    }

    [Fact]
    public void SetWorkbookProperties_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_set_workbook.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_workbook_properties", sessionId: sessionId, title: "Updated Title",
            author: "Updated Author");
        Assert.Contains("Workbook properties updated", result);
        Assert.Contains("session", result); // Verify session was used
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Updated Title", workbook.BuiltInDocumentProperties.Title);
        Assert.Equal("Updated Author", workbook.BuiltInDocumentProperties.Author);
    }

    [Fact]
    public void GetSheetProperties_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_sheet_properties", sessionId: sessionId, sheetIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("name", out _));
        Assert.Equal(5, json.RootElement.GetProperty("dataRowCount").GetInt32());
    }

    [Fact]
    public void EditSheetProperties_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit_sheet_properties", sessionId: sessionId, sheetIndex: 0, name: "RenamedSheet");
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("RenamedSheet", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void GetSheetInfo_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_info.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("TestSheet");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_sheet_info", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() >= 2);
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
        Assert.Contains("Session Title", result);
    }

    #endregion
}