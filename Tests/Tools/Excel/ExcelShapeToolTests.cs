using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Shape;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelShapeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelShapeToolTests : ExcelTestBase
{
    private readonly ExcelShapeTool _tool;

    public ExcelShapeToolTests()
    {
        _tool = new ExcelShapeTool(SessionManager);
    }

    private string CreateWorkbookWithShape(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var sheet = workbook.Worksheets[0];
        sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 1, 0, 1, 0, 100, 200);
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddShape()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, shapeType: "Rectangle", upperLeftRow: 1,
            upperLeftColumn: 1, width: 100, height: 50, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Shapes.Count >= 1);
    }

    [Fact]
    public void Get_ShouldReturnShapes()
    {
        var workbookPath = CreateWorkbookWithShape("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetShapesExcelResult>(result);
        Assert.True(data.Count >= 1);
        Assert.True(data.Items.Count >= 1);
    }

    [Fact]
    public void Edit_ShouldModifyShape()
    {
        var workbookPath = CreateWorkbookWithShape("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, shapeIndex: 0, text: "Updated", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("updated", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Updated", workbook.Worksheets[0].Shapes[0].Text);
    }

    [Fact]
    public void Delete_ShouldDeleteShape()
    {
        var workbookPath = CreateWorkbookWithShape("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, shapeIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Shapes);
    }

    [Fact]
    public void AddTextBox_ShouldAddTextBox()
    {
        var workbookPath = CreateExcelWorkbook("test_add_textbox.xlsx");
        var outputPath = CreateTestFilePath("test_add_textbox_output.xlsx");
        var result = _tool.Execute("add_textbox", workbookPath, text: "Hello", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Shapes.Count >= 1);
        Assert.Equal("Hello", workbook.Worksheets[0].Shapes[0].Text);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, shapeType: "Rectangle", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message, StringComparison.OrdinalIgnoreCase);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSession_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, shapeType: "Rectangle",
            upperLeftRow: 1, upperLeftColumn: 1, width: 100, height: 50);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Shapes.Count >= 1);
    }

    [Fact]
    public void Get_WithSession_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithShape("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetShapesExcelResult>(result);
        Assert.True(data.Count >= 1);
        var output = GetResultOutput<GetShapesExcelResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Edit_WithSession_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithShape("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, shapeIndex: 0, text: "SessionText");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionText", workbook.Worksheets[0].Shapes[0].Text);
    }

    [Fact]
    public void Delete_WithSession_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithShape("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, shapeIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Shapes);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithShape("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var data = GetResultData<GetShapesExcelResult>(result);
        Assert.True(data.Count >= 1);
    }

    #endregion
}
