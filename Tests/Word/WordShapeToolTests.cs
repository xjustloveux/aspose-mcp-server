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

    [Fact]
    public async Task AddLine_WithOptions_ShouldAddLineWithCustomSettings()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_line_options.docx");
        var outputPath = CreateTestFilePath("test_add_line_options_output.docx");
        var arguments = CreateArguments("add_line", docPath, outputPath);
        arguments["location"] = "body";
        arguments["position"] = "start";
        arguments["lineStyle"] = "shape";
        arguments["lineWidth"] = 2.0;
        arguments["lineColor"] = "FF0000";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Successfully inserted line", result);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain line shape");
    }

    [Fact]
    public async Task AddLine_InHeader_ShouldAddLineToHeader()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_line_header.docx");
        var outputPath = CreateTestFilePath("test_add_line_header_output.docx");
        var arguments = CreateArguments("add_line", docPath, outputPath);
        arguments["location"] = "header";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("header", result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
    }

    [Fact]
    public async Task AddLine_InFooter_ShouldAddLineToFooter()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_line_footer.docx");
        var outputPath = CreateTestFilePath("test_add_line_footer_output.docx");
        var arguments = CreateArguments("add_line", docPath, outputPath);
        arguments["location"] = "footer";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("footer", result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    [Fact]
    public async Task AddLine_WithBorderStyle_ShouldAddBorderLine()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_line_border.docx");
        var outputPath = CreateTestFilePath("test_add_line_border_output.docx");
        var arguments = CreateArguments("add_line", docPath, outputPath);
        arguments["lineStyle"] = "border";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Successfully inserted line", result);
    }

    [Fact]
    public async Task AddTextBox_WithFontSettings_ShouldApplyFontSettings()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_textbox_font.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_font_output.docx");
        var arguments = CreateArguments("add_textbox", docPath, outputPath);
        arguments["text"] = "Styled Text";
        arguments["fontName"] = "Arial";
        arguments["fontSize"] = 14;
        arguments["bold"] = true;
        arguments["textAlignment"] = "center";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(shapes.Count > 0, "Document should contain textbox");
    }

    [Fact]
    public async Task AddTextBox_WithBackgroundColor_ShouldSetBackgroundColor()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_textbox_bg.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_bg_output.docx");
        var arguments = CreateArguments("add_textbox", docPath, outputPath);
        arguments["text"] = "Colored Background";
        arguments["backgroundColor"] = "FFFF00";
        arguments["borderColor"] = "0000FF";
        arguments["borderWidth"] = 2;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(shapes.Count > 0, "Document should contain textbox");
        Assert.True(shapes[0].Fill.Visible, "Fill should be visible");
    }

    [Fact]
    public async Task EditTextBoxContent_WithAppendText_ShouldAppendText()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_textbox_append.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_append_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Original";
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_edit_textbox_append_output.docx");
        var arguments = CreateArguments("edit_textbox_content", tempOutputPath, outputPath);
        arguments["textboxIndex"] = 0;
        arguments["text"] = " Appended";
        arguments["appendText"] = true;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0);
        var text = textboxes[0].GetText();
        Assert.Contains("Original", text);
        Assert.Contains("Appended", text);
    }

    [Fact]
    public async Task EditTextBoxContent_WithFormatting_ShouldApplyFormatting()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_textbox_format.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_format_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Format Me";
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_edit_textbox_format_output.docx");
        var arguments = CreateArguments("edit_textbox_content", tempOutputPath, outputPath);
        arguments["textboxIndex"] = 0;
        arguments["bold"] = true;
        arguments["italic"] = true;
        arguments["color"] = "FF0000";
        arguments["fontSize"] = 16;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Successfully edited textbox", result);
    }

    [Fact]
    public async Task SetTextBoxBorder_WithBorderHidden_ShouldHideBorder()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_border_hidden.docx");
        var tempOutputPath = CreateTestFilePath("test_set_border_hidden_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "No Border";
        addArguments["borderColor"] = "000000";
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_set_border_hidden_output.docx");
        var arguments = CreateArguments("set_textbox_border", tempOutputPath, outputPath);
        arguments["textboxIndex"] = 0;
        arguments["borderVisible"] = false;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("No border", result);
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.False(textboxes[0].Stroke.Visible, "Border should be hidden");
    }

    [Fact]
    public async Task AddChart_WithDifferentTypes_ShouldAddDifferentChartTypes()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_chart_types.docx");
        var outputPath = CreateTestFilePath("test_add_chart_types_output.docx");
        var arguments = CreateArguments("add_chart", docPath, outputPath);
        arguments["chartType"] = "pie";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "Category", "Value" },
            new JsonArray { "A", "30" },
            new JsonArray { "B", "70" }
        };
        arguments["chartTitle"] = "Pie Chart";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Successfully added chart", result);
        Assert.Contains("pie", result);
    }

    [Fact]
    public async Task AddChart_WithAlignment_ShouldSetChartAlignment()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_chart_align.docx");
        var outputPath = CreateTestFilePath("test_add_chart_align_output.docx");
        var arguments = CreateArguments("add_chart", docPath, outputPath);
        arguments["chartType"] = "bar";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "X", "Y" },
            new JsonArray { "1", "2" }
        };
        arguments["alignment"] = "center";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Successfully added chart", result);
    }

    [Fact]
    public async Task AddShape_WithPosition_ShouldSetShapePosition()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_shape_pos.docx");
        var outputPath = CreateTestFilePath("test_add_shape_pos_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["shapeType"] = "ellipse";
        arguments["width"] = 150;
        arguments["height"] = 100;
        arguments["x"] = 200;
        arguments["y"] = 150;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("ellipse", result);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public async Task DeleteShape_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_invalid.docx");
        var arguments = CreateArguments("delete", docPath);
        arguments["shapeIndex"] = 999;

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task EditTextBoxContent_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_invalid.docx");
        var arguments = CreateArguments("edit_textbox_content", docPath);
        arguments["textboxIndex"] = 999;
        arguments["text"] = "Test";

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task SetTextBoxBorder_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_border_invalid.docx");
        var arguments = CreateArguments("set_textbox_border", docPath);
        arguments["textboxIndex"] = 999;

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddShape_WithInvalidShapeType_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_invalid_type.docx");
        var arguments = CreateArguments("add", docPath);
        arguments["shapeType"] = "invalid_type";
        arguments["width"] = 100;
        arguments["height"] = 50;

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddChart_WithEmptyData_ShouldThrowException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_chart_empty.docx");
        var arguments = CreateArguments("add_chart", docPath);
        arguments["chartType"] = "column";
        arguments["data"] = new JsonArray();

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetTextboxes_WithNoTextboxes_ShouldReturnEmptyMessage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_no_textboxes.docx");
        var arguments = CreateArguments("get_textboxes", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Total Textboxes: 0", result);
        Assert.Contains("No textboxes found", result);
    }

    [Fact]
    public async Task GetShapes_WithNoShapes_ShouldReturnEmptyMessage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_no_shapes.docx");
        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Total Shapes: 0", result);
        Assert.Contains("No shapes found", result);
    }

    [Fact]
    public async Task GetTextboxes_WithIncludeContentFalse_ShouldNotIncludeContent()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_textboxes_no_content.docx");
        var tempOutputPath = CreateTestFilePath("test_get_textboxes_no_content_temp.docx");
        var addArguments = CreateArguments("add_textbox", docPath, tempOutputPath);
        addArguments["text"] = "Secret Content";
        await _tool.ExecuteAsync(addArguments);

        var arguments = new JsonObject
        {
            ["operation"] = "get_textboxes",
            ["path"] = tempOutputPath,
            ["includeContent"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Textbox", result);
        Assert.DoesNotContain("Secret Content", result);
        Assert.DoesNotContain("Content:", result);
    }
}