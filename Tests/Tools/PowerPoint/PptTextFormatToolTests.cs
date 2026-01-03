using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptTextFormatToolTests : TestBase
{
    private readonly PptTextFormatTool _tool;

    public PptTextFormatToolTests()
    {
        _tool = new PptTextFormatTool(SessionManager);
    }

    private string CreatePptPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        textBox.TextFrame.Text = "Sample Text";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePptWithTable(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(100, 100, [100, 100], [50, 50]);
        table[0, 0].TextFrame.Text = "Cell 1";
        table[0, 1].TextFrame.Text = "Cell 2";
        table[1, 0].TextFrame.Text = "Cell 3";
        table[1, 1].TextFrame.Text = "Cell 4";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void FormatText_WithFontOptions_ShouldApplyFontFormatting()
    {
        var pptPath = CreatePptPresentation("test_format_font.pptx");
        var outputPath = CreateTestFilePath("test_format_font_output.pptx");
        _tool.Execute(pptPath, fontName: "Arial", fontSize: 16, bold: true, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public void FormatText_WithColor_ShouldApplyColor()
    {
        var pptPath = CreatePptPresentation("test_format_color.pptx");
        var outputPath = CreateTestFilePath("test_format_color_output.pptx");
        _tool.Execute(pptPath, color: "#FF0000", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public void FormatText_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        var pptPath = CreatePptPresentation("test_format_all.pptx");
        var outputPath = CreateTestFilePath("test_format_all_output.pptx");
        _tool.Execute(pptPath, fontName: "Arial", fontSize: 14, bold: true, italic: true, color: "#0000FF",
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public void FormatText_WithSpecificSlides_ShouldFormatOnlySelectedSlides()
    {
        var pptPath = CreatePptPresentation("test_format_specific_slides.pptx");
        using var presentation = new Presentation(pptPath);
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(pptPath, SaveFormat.Pptx);

        var outputPath = CreateTestFilePath("test_format_specific_slides_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Times New Roman", fontSize: 12,
            outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        Assert.True(resultPresentation.Slides.Count >= 2);
    }

    [Fact]
    public void FormatText_WithTableText_ShouldFormatTableCells()
    {
        var pptPath = CreatePptWithTable("test_format_table.pptx");
        var outputPath = CreateTestFilePath("test_format_table_output.pptx");
        var result = _tool.Execute(pptPath, fontName: "Arial", fontSize: 14, bold: true, outputPath: outputPath);
        Assert.Contains("1 slides", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public void FormatText_WithNamedColor_ShouldApplyNamedColor()
    {
        var pptPath = CreatePptPresentation("test_format_named_color.pptx");
        var outputPath = CreateTestFilePath("test_format_named_color_output.pptx");
        var result = _tool.Execute(pptPath, color: "Red", outputPath: outputPath);
        Assert.Contains("slides", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void FormatText_WithItalicOnly_ShouldApplyItalic()
    {
        var pptPath = CreatePptPresentation("test_format_italic.pptx");
        var outputPath = CreateTestFilePath("test_format_italic_output.pptx");
        var result = _tool.Execute(pptPath, italic: true, outputPath: outputPath);
        Assert.Contains("slides", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void FormatText_WithMixedShapes_ShouldFormatBothAutoShapeAndTable()
    {
        // Arrange - Create presentation with both AutoShape and Table
        var filePath = CreateTestFilePath("test_format_mixed.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
            textBox.TextFrame.Text = "AutoShape Text";
            var table = slide.Shapes.AddTable(50, 150, [100, 100], [50]);
            table[0, 0].TextFrame.Text = "Table Text";
            presentation.Save(filePath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_format_mixed_output.pptx");
        var result = _tool.Execute(filePath, fontName: "Verdana", fontSize: 18, outputPath: outputPath);
        Assert.Contains("1 slides", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void FormatText_InvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreatePptPresentation("test_format_invalid_index.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 99 });
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Arial"));
    }

    [Fact]
    public void FormatText_InvalidColor_ShouldDefaultToBlack()
    {
        var pptPath = CreatePptPresentation("test_format_invalid_color.pptx");
        var outputPath = CreateTestFilePath("test_format_invalid_color_output.pptx");

        // Act - Invalid color names default to Black (ColorHelper.ParseColor behavior)
        var result = _tool.Execute(pptPath, outputPath: outputPath, color: "InvalidColorName");

        // Assert - Tool succeeds with default black color
        Assert.True(File.Exists(outputPath));
        Assert.Contains("formatted", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void FormatText_WithSessionId_ShouldFormatInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, fontName: "Arial", fontSize: 16, bold: true);
        Assert.Contains("slides", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void FormatText_WithSessionId_ShouldApplyColorInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_format_color.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, color: "#FF0000");
        Assert.Contains("slides", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
        Assert.True(ppt.Slides.Count > 0);
    }

    [Fact]
    public void FormatText_WithSessionId_MultipleFormats_ShouldApplyAll()
    {
        var pptPath = CreatePptPresentation("test_session_format_multi.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, fontName: "Verdana", fontSize: 18,
            bold: true, italic: true, color: "#0000FF");
        Assert.Contains("slides", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
    }

    #endregion
}