using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Shape;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordShapeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordShapeToolTests : WordTestBase
{
    private readonly WordShapeTool _tool;

    public WordShapeToolTests()
    {
        _tool = new WordShapeTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddShapeAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_shape.docx");
        var outputPath = CreateTestFilePath("test_add_shape_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, shapeType: "Rectangle", width: 100, height: 50);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void Get_ShouldReturnShapesFromFile()
    {
        var docPath = CreateWordDocument("test_get_shapes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetShapesWordResult>(result);
        Assert.Contains("Shape", data.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_ShouldDeleteShapeAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_shape.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_shape_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, shapeIndex: 0);
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true);
        Assert.Equal(0, shapes.Count);
    }

    [Fact]
    public void AddLine_ShouldAddLineAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_line.docx");
        var outputPath = CreateTestFilePath("test_add_line_output.docx");
        var result = _tool.Execute("add_line", docPath, outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void AddTextBox_ShouldAddTextBoxAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_textbox.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_output.docx");
        var result = _tool.Execute("add_textbox", docPath, outputPath: outputPath, text: "Test TextBox");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0);
    }

    [Fact]
    public void AddChart_ShouldAddChartAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_chart.docx");
        var outputPath = CreateTestFilePath("test_add_chart_output.docx");
        var chartData = new[] { new[] { "A", "B" }, new[] { "1", "2" } };
        var result = _tool.Execute("add_chart", docPath, outputPath: outputPath, chartType: "column", data: chartData);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        _tool.Execute(operation, docPath, outputPath: outputPath, shapeType: "Rectangle", width: 100, height: 50);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
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
    public void Get_WithSessionId_ShouldReturnShapes()
    {
        var docPath = CreateWordDocument("test_session_get_shapes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetShapesWordResult>(result);
        Assert.Contains("Shape", data.Content, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<GetShapesWordResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddShapeInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_shape.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, shapeType: "Rectangle", width: 100, height: 50);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_shape.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, shapeIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.Equal(0, shapes.Count);
    }

    [Fact]
    public void AddLine_WithSessionId_ShouldAddLineInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_line.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_line", sessionId: sessionId);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void AddTextBox_WithSessionId_ShouldAddTextBoxInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_textbox.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_textbox", sessionId: sessionId, text: "Session TextBox",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var textboxes = sessionDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0);
    }

    [Fact]
    public void AddChart_WithSessionId_ShouldAddChartInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_chart.docx");
        var sessionId = OpenSession(docPath);
        var chartData = new[] { new[] { "A", "B" }, new[] { "1", "2" } };

        var result = _tool.Execute("add_chart", sessionId: sessionId, chartType: "column", data: chartData);

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_shape.docx");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        builder1.InsertShape(ShapeType.Rectangle, 100, 50);
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_shape.docx");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        builder2.InsertShape(ShapeType.Ellipse, 80, 80);
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);

        var data = GetResultData<GetShapesWordResult>(result);
        Assert.Contains("Ellipse", data.Content, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<GetShapesWordResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
