using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordShapeToolTests : WordTestBase
{
    private readonly WordShapeTool _tool;

    public WordShapeToolTests()
    {
        _tool = new WordShapeTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddShape_ShouldAddShape()
    {
        var docPath = CreateWordDocument("test_add_shape.docx");
        var outputPath = CreateTestFilePath("test_add_shape_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, shapeType: "Rectangle", width: 100, height: 50);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one shape");
    }

    [Fact]
    public void GetShapes_ShouldReturnAllShapes()
    {
        var docPath = CreateWordDocument("test_get_shapes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Shape", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteShape_ShouldDeleteShape()
    {
        var docPath = CreateWordDocument("test_delete_shape.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var shapesBefore = doc.GetChildNodes(NodeType.Shape, true).Count;
        Assert.True(shapesBefore > 0, "Shape should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_shape_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, shapeIndex: 0);
        var resultDoc = new Document(outputPath);
        var shapesAfter = resultDoc.GetChildNodes(NodeType.Shape, true).Count;
        Assert.True(shapesAfter < shapesBefore,
            $"Shape should be deleted. Before: {shapesBefore}, After: {shapesAfter}");
    }

    [Fact]
    public void AddLine_ShouldAddLineShape()
    {
        var docPath = CreateWordDocument("test_add_line.docx");
        var outputPath = CreateTestFilePath("test_add_line_output.docx");
        _tool.Execute("add_line", docPath, outputPath: outputPath);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one line shape");
    }

    [Fact]
    public void AddTextBox_ShouldAddTextBox()
    {
        var docPath = CreateWordDocument("test_add_textbox.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_output.docx");
        _tool.Execute("add_textbox", docPath, outputPath: outputPath, text: "TextBox Content",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0, "Document should contain at least one textbox");
        Assert.True(shapes.Any(s => s.GetText().Contains("TextBox Content")),
            "TextBox should contain 'TextBox Content'");
    }

    [Fact]
    public void GetTextboxes_ShouldReturnAllTextboxes()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_get_textboxes.docx");
        var tempOutputPath = CreateTestFilePath("test_get_textboxes_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Test TextBox",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);
        var result = _tool.Execute("get_textboxes", tempOutputPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("TextBox", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditTextBoxContent_ShouldEditTextBoxContent()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_edit_textbox_content.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_content_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Original Text",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);

        var outputPath = CreateTestFilePath("test_edit_textbox_content_output.docx");
        _tool.Execute("edit_textbox_content", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            text: "Updated TextBox Content");
        var resultDoc = new Document(outputPath);
        var textboxes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0, "Document should contain textbox after editing");
        Assert.Contains("Updated TextBox Content", textboxes[0].GetText());
    }

    [Fact]
    public void SetTextBoxBorder_ShouldSetTextBoxBorder()
    {
        // Arrange - Create textbox using the tool's add_textbox operation
        var docPath = CreateWordDocument("test_set_textbox_border.docx");
        var tempOutputPath = CreateTestFilePath("test_set_textbox_border_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Test TextBox",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);

        var outputPath = CreateTestFilePath("test_set_textbox_border_output.docx");
        _tool.Execute("set_textbox_border", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            borderColor: "FF0000", borderWidth: 2);
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain textbox after setting border");
    }

    [Fact]
    public void AddChart_ShouldAddChart()
    {
        var docPath = CreateWordDocument("test_add_chart.docx");
        var outputPath = CreateTestFilePath("test_add_chart_output.docx");
        var data = new[]
        {
            new[] { "A", "B" },
            new[] { "1", "2" }
        };
        _tool.Execute("add_chart", docPath, outputPath: outputPath, chartType: "Column", data: data);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain at least one chart");
    }

    [Fact]
    public void AddLine_WithOptions_ShouldAddLineWithCustomSettings()
    {
        var docPath = CreateWordDocument("test_add_line_options.docx");
        var outputPath = CreateTestFilePath("test_add_line_options_output.docx");
        var result = _tool.Execute("add_line", docPath, outputPath: outputPath,
            location: "body", position: "start", lineStyle: "shape", lineWidth: 2.0, lineColor: "FF0000");
        Assert.Contains("Successfully inserted line", result);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0, "Document should contain line shape");
    }

    [Fact]
    public void AddLine_InHeader_ShouldAddLineToHeader()
    {
        var docPath = CreateWordDocument("test_add_line_header.docx");
        var outputPath = CreateTestFilePath("test_add_line_header_output.docx");
        var result = _tool.Execute("add_line", docPath, outputPath: outputPath, location: "header");
        Assert.Contains("header", result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
    }

    [Fact]
    public void AddLine_InFooter_ShouldAddLineToFooter()
    {
        var docPath = CreateWordDocument("test_add_line_footer.docx");
        var outputPath = CreateTestFilePath("test_add_line_footer_output.docx");
        var result = _tool.Execute("add_line", docPath, outputPath: outputPath, location: "footer");
        Assert.Contains("footer", result);
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    [Fact]
    public void AddLine_WithBorderStyle_ShouldAddBorderLine()
    {
        var docPath = CreateWordDocument("test_add_line_border.docx");
        var outputPath = CreateTestFilePath("test_add_line_border_output.docx");
        var result = _tool.Execute("add_line", docPath, outputPath: outputPath, lineStyle: "border");
        Assert.Contains("Successfully inserted line", result);
    }

    [Fact]
    public void AddTextBox_WithFontSettings_ShouldApplyFontSettings()
    {
        var docPath = CreateWordDocument("test_add_textbox_font.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_font_output.docx");
        _tool.Execute("add_textbox", docPath, outputPath: outputPath, text: "Styled Text",
            fontName: "Arial", fontSize: 14, bold: true, textAlignment: "center");
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(shapes.Count > 0, "Document should contain textbox");
    }

    [Fact]
    public void AddTextBox_WithBackgroundColor_ShouldSetBackgroundColor()
    {
        var docPath = CreateWordDocument("test_add_textbox_bg.docx");
        var outputPath = CreateTestFilePath("test_add_textbox_bg_output.docx");
        _tool.Execute("add_textbox", docPath, outputPath: outputPath, text: "Colored Background",
            backgroundColor: "FFFF00", borderColor: "0000FF", borderWidth: 2);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(shapes.Count > 0, "Document should contain textbox");
        Assert.True(shapes[0].Fill.Visible, "Fill should be visible");
    }

    [Fact]
    public void EditTextBoxContent_WithAppendText_ShouldAppendText()
    {
        var docPath = CreateWordDocument("test_edit_textbox_append.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_append_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Original");

        var outputPath = CreateTestFilePath("test_edit_textbox_append_output.docx");
        _tool.Execute("edit_textbox_content", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            text: " Appended", appendText: true);
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0);
        var text = textboxes[0].GetText();
        Assert.Contains("Original", text);
        Assert.Contains("Appended", text);
    }

    [Fact]
    public void EditTextBoxContent_WithFormatting_ShouldApplyFormatting()
    {
        var docPath = CreateWordDocument("test_edit_textbox_format.docx");
        var tempOutputPath = CreateTestFilePath("test_edit_textbox_format_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Format Me");

        var outputPath = CreateTestFilePath("test_edit_textbox_format_output.docx");
        var result = _tool.Execute("edit_textbox_content", tempOutputPath, outputPath: outputPath,
            textboxIndex: 0, bold: true, italic: true, color: "FF0000", fontSize: 16);
        Assert.Contains("Successfully edited textbox", result);
    }

    [Fact]
    public void SetTextBoxBorder_WithBorderHidden_ShouldHideBorder()
    {
        var docPath = CreateWordDocument("test_set_border_hidden.docx");
        var tempOutputPath = CreateTestFilePath("test_set_border_hidden_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "No Border", borderColor: "000000");

        var outputPath = CreateTestFilePath("test_set_border_hidden_output.docx");
        var result = _tool.Execute("set_textbox_border", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            borderVisible: false);
        Assert.Contains("No border", result);
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.False(textboxes[0].Stroke.Visible, "Border should be hidden");
    }

    [Theory]
    [InlineData("solid", DashStyle.Solid)]
    [InlineData("dash", DashStyle.Dash)]
    [InlineData("dot", DashStyle.Dot)]
    [InlineData("dashdot", DashStyle.DashDot)]
    [InlineData("dashdotdot", DashStyle.LongDashDotDot)]
    [InlineData("rounddot", DashStyle.ShortDot)]
    public void SetTextBoxBorder_WithBorderStyle_ShouldApplyCorrectDashStyle(string borderStyle,
        DashStyle expectedDashStyle)
    {
        var docPath = CreateWordDocument($"test_border_style_{borderStyle}.docx");
        var tempOutputPath = CreateTestFilePath($"test_border_style_{borderStyle}_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Styled Border");

        var outputPath = CreateTestFilePath($"test_border_style_{borderStyle}_output.docx");
        var result = _tool.Execute("set_textbox_border", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            borderVisible: true, borderColor: "0000FF", borderWidth: 2, borderStyle: borderStyle);

        Assert.Contains($"Style: {borderStyle}", result);
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes[0].Stroke.Visible, "Border should be visible");
        Assert.Equal(expectedDashStyle, textboxes[0].Stroke.DashStyle);
    }

    [Fact]
    public void SetTextBoxBorder_WithUnknownBorderStyle_ShouldDefaultToSolid()
    {
        var docPath = CreateWordDocument("test_border_style_unknown.docx");
        var tempOutputPath = CreateTestFilePath("test_border_style_unknown_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Unknown Style");

        var outputPath = CreateTestFilePath("test_border_style_unknown_output.docx");
        var result = _tool.Execute("set_textbox_border", tempOutputPath, outputPath: outputPath, textboxIndex: 0,
            borderVisible: true, borderColor: "FF0000", borderWidth: 1, borderStyle: "unknown_style");

        Assert.Contains("Style: unknown_style", result);
        var doc = new Document(outputPath);
        var textboxes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.Equal(DashStyle.Solid, textboxes[0].Stroke.DashStyle);
    }

    [Fact]
    public void AddChart_WithDifferentTypes_ShouldAddDifferentChartTypes()
    {
        var docPath = CreateWordDocument("test_add_chart_types.docx");
        var outputPath = CreateTestFilePath("test_add_chart_types_output.docx");
        var data = new[]
        {
            new[] { "Category", "Value" },
            new[] { "A", "30" },
            new[] { "B", "70" }
        };
        var result = _tool.Execute("add_chart", docPath, outputPath: outputPath, chartType: "pie", data: data,
            chartTitle: "Pie Chart");
        Assert.Contains("Successfully added chart", result);
        Assert.Contains("pie", result);
    }

    [Fact]
    public void AddChart_WithAlignment_ShouldSetChartAlignment()
    {
        var docPath = CreateWordDocument("test_add_chart_align.docx");
        var outputPath = CreateTestFilePath("test_add_chart_align_output.docx");
        var data = new[]
        {
            new[] { "X", "Y" },
            new[] { "1", "2" }
        };
        var result = _tool.Execute("add_chart", docPath, outputPath: outputPath, chartType: "bar", data: data,
            alignment: "center");
        Assert.Contains("Successfully added chart", result);
    }

    [Fact]
    public void AddShape_WithPosition_ShouldSetShapePosition()
    {
        var docPath = CreateWordDocument("test_add_shape_pos.docx");
        var outputPath = CreateTestFilePath("test_add_shape_pos_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, shapeType: "ellipse", width: 150,
            height: 100, x: 200, y: 150);
        Assert.Contains("ellipse", result);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void DeleteShape_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_delete_invalid.docx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("delete", docPath, shapeIndex: 999));
    }

    [Fact]
    public void EditTextBoxContent_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_edit_invalid.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_textbox_content", docPath, textboxIndex: 999, text: "Test"));
    }

    [Fact]
    public void SetTextBoxBorder_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_border_invalid.docx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("set_textbox_border", docPath, textboxIndex: 999));
    }

    [Fact]
    public void AddShape_WithInvalidShapeType_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_add_invalid_type.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, shapeType: "invalid_type", width: 100, height: 50));
    }

    [Fact]
    public void AddChart_WithEmptyData_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_add_chart_empty.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_chart", docPath, chartType: "column", data: Array.Empty<string[]>()));
    }

    [Fact]
    public void GetTextboxes_WithNoTextboxes_ShouldReturnEmptyMessage()
    {
        var docPath = CreateWordDocument("test_get_no_textboxes.docx");
        var result = _tool.Execute("get_textboxes", docPath);
        Assert.Contains("Total Textboxes: 0", result);
        Assert.Contains("No textboxes found", result);
    }

    [Fact]
    public void GetShapes_WithNoShapes_ShouldReturnEmptyMessage()
    {
        var docPath = CreateWordDocument("test_get_no_shapes.docx");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("Total Shapes: 0", result);
        Assert.Contains("No shapes found", result);
    }

    [Fact]
    public void GetTextboxes_WithIncludeContentFalse_ShouldNotIncludeContent()
    {
        var docPath = CreateWordDocument("test_get_textboxes_no_content.docx");
        var tempOutputPath = CreateTestFilePath("test_get_textboxes_no_content_temp.docx");
        _tool.Execute("add_textbox", docPath, outputPath: tempOutputPath, text: "Secret Content");
        var result = _tool.Execute("get_textboxes", tempOutputPath, includeContent: false);
        Assert.Contains("Textbox", result);
        Assert.DoesNotContain("Secret Content", result);
        Assert.DoesNotContain("Content:", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void AddShape_WithMissingShapeType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_missing_shape_type.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, shapeType: "", width: 100, height: 50));

        Assert.Contains("shapeType", ex.Message);
    }

    [Fact]
    public void AddChart_WithEmptyChartType_ShouldDefaultToColumn()
    {
        var docPath = CreateWordDocument("test_missing_chart_type.docx");
        var outputPath = CreateTestFilePath("test_missing_chart_type_output.docx");
        var data = new[] { new[] { "A", "B" }, new[] { "1", "2" } };

        // Act - Empty chartType defaults to "Column" (switch default case)
        var result = _tool.Execute("add_chart", docPath, outputPath: outputPath, chartType: "", data: data);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Successfully added chart", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetShapes_WithSessionId_ShouldReturnShapes()
    {
        var docPath = CreateWordDocument("test_session_get_shapes.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Shape", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddShape_WithSessionId_ShouldAddShapeInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_shape.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, shapeType: "Rectangle", width: 100, height: 50);
        Assert.Contains("Rectangle", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the shape
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void DeleteShape_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_shape.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, shapeIndex: 0);

        // Verify in-memory document has no shapes
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true);
        Assert.Equal(0, shapes.Count);
    }

    [Fact]
    public void AddTextBox_WithSessionId_ShouldAddTextBoxInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_textbox.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_textbox", sessionId: sessionId, text: "Session TextBox",
            positionX: 100, positionY: 100, textboxWidth: 200, textboxHeight: 100);
        Assert.Contains("TextBox", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the textbox
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var textboxes = sessionDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        Assert.True(textboxes.Count > 0);
        Assert.Contains("Session TextBox", textboxes[0].GetText());
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId (which has Ellipse, not Rectangle)
        Assert.Contains("Ellipse", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}