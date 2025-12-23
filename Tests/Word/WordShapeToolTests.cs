using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordShapeToolTests : WordTestBase
{
    private readonly WordShapeTool _tool = new();

    [Fact]
    public async Task AddShape_ShouldAddShape()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_shape.docx");
        var outputPath = CreateTestFilePath("test_add_shape_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["shapeType"] = "Rectangle";
        arguments["width"] = 100;
        arguments["height"] = 50;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one shape");
    }

    [Fact]
    public async Task GetShapes_ShouldReturnAllShapes()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_shapes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Shape", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteShape_ShouldDeleteShape()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_shape.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var shapesBefore = doc.GetChildNodes(NodeType.Shape, true).Count;
        Assert.True(shapesBefore > 0, "Shape should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_shape_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["shapeIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapesAfter = resultDoc.GetChildNodes(NodeType.Shape, true).Count;
        Assert.True(shapesAfter < shapesBefore,
            $"Shape should be deleted. Before: {shapesBefore}, After: {shapesAfter}");
    }

    [Fact]
    public async Task AddLine_ShouldAddLineShape()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_line.docx");
        var outputPath = CreateTestFilePath("test_add_line_output.docx");
        var arguments = CreateArguments("add_line", docPath, outputPath);
        arguments["x1"] = 100;
        arguments["y1"] = 100;
        arguments["x2"] = 200;
        arguments["y2"] = 200;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one line shape");
    }

    [Fact]
    public async Task AddTextBox_ShouldAddTextBox()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_textbox.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_output.docx");
        var arguments = CreateArguments("add_textbox", docPath, outputPath);
        arguments["text"] = "TextBox Content";
        arguments["x"] = 100;
        arguments["y"] = 100;
        arguments["width"] = 200;
        arguments["height"] = 100;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one textbox");
    }

    [Fact]
    public async Task GetTextboxes_ShouldReturnAllTextboxes()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_get_textboxes.docx");
        var tempOutputPath = CreateTestFilePath("test_get_textboxes_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Test TextBox";
        addArguments["x"] = 100;
        addArguments["y"] = 100;
        addArguments["width"] = 200;
        addArguments["height"] = 100;
        await _tool.ExecuteAsync(addArguments);

        var arguments = new JsonObject
        {
            ["operation"] = "get_textboxes",
            ["path"] = tempOutputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("TextBox", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task EditTextBoxContent_ShouldEditTextBoxContent()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_edit_textbox_content.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_content_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Original Text";
        addArguments["x"] = 100;
        addArguments["y"] = 100;
        addArguments["width"] = 200;
        addArguments["height"] = 100;
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_edit_textbox_content_output.docx");
        var arguments = CreateArguments("edit_textbox_content", tempOutputPath, outputPath);
        arguments["textboxIndex"] = 0;
        arguments["text"] = "Updated TextBox Content";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain textbox after editing");
    }

    [Fact]
    public async Task SetTextBoxBorder_ShouldSetTextBoxBorder()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_set_textbox_border.docx");
        var tempOutputPath = CreateTestFilePath("test_set_textbox_border_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Test TextBox";
        addArguments["x"] = 100;
        addArguments["y"] = 100;
        addArguments["width"] = 200;
        addArguments["height"] = 100;
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_set_textbox_border_output.docx");
        var arguments = CreateArguments("set_textbox_border", tempOutputPath, outputPath);
        arguments["textboxIndex"] = 0;
        arguments["color"] = "#FF0000";
        arguments["width"] = 2;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain textbox after setting border");
    }

    [Fact]
    public async Task AddChart_ShouldAddChart()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_chart.docx");
        var outputPath = CreateTestFilePath("test_add_chart_output.docx");
        var arguments = CreateArguments("add_chart", docPath, outputPath);
        arguments["chartType"] = "Column";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "A", "B" },
            new JsonArray { "1", "2" }
        };
        arguments["x"] = 100;
        arguments["y"] = 100;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one chart");
    }
}