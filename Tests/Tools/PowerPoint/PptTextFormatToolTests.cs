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

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        textBox.TextFrame.Text = "Sample Text";
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithTable(string fileName)
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

    #region General

    [Fact]
    public void Execute_WithFontOptions_ShouldApplyFontFormatting()
    {
        var pptPath = CreateTestPresentation("test_format_font.pptx");
        var outputPath = CreateTestFilePath("test_format_font_output.pptx");
        var result = _tool.Execute(pptPath, fontName: "Arial", fontSize: 16, bold: true, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithHexColor_ShouldApplyColor()
    {
        var pptPath = CreateTestPresentation("test_format_color.pptx");
        var outputPath = CreateTestFilePath("test_format_color_output.pptx");
        var result = _tool.Execute(pptPath, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithNamedColor_ShouldApplyNamedColor()
    {
        var pptPath = CreateTestPresentation("test_format_named_color.pptx");
        var outputPath = CreateTestFilePath("test_format_named_color_output.pptx");
        var result = _tool.Execute(pptPath, color: "Red", outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        var pptPath = CreateTestPresentation("test_format_all.pptx");
        var outputPath = CreateTestFilePath("test_format_all_output.pptx");
        var result = _tool.Execute(pptPath, fontName: "Arial", fontSize: 14, bold: true, italic: true, color: "#0000FF",
            outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithItalicOnly_ShouldApplyItalic()
    {
        var pptPath = CreateTestPresentation("test_format_italic.pptx");
        var outputPath = CreateTestFilePath("test_format_italic_output.pptx");
        var result = _tool.Execute(pptPath, italic: true, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithBoldOnly_ShouldApplyBold()
    {
        var pptPath = CreateTestPresentation("test_format_bold.pptx");
        var outputPath = CreateTestFilePath("test_format_bold_output.pptx");
        var result = _tool.Execute(pptPath, bold: true, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithFontSizeOnly_ShouldApplyFontSize()
    {
        var pptPath = CreateTestPresentation("test_format_size.pptx");
        var outputPath = CreateTestFilePath("test_format_size_output.pptx");
        var result = _tool.Execute(pptPath, fontSize: 24, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithSpecificSlides_ShouldFormatOnlySelectedSlides()
    {
        var pptPath = CreateTestPresentation("test_format_specific.pptx", 3);
        var outputPath = CreateTestFilePath("test_format_specific_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Times New Roman", fontSize: 12,
            outputPath: outputPath);
        Assert.Contains("1 slides", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Execute_WithMultipleSlideIndices_ShouldFormatMultipleSlides()
    {
        var pptPath = CreateTestPresentation("test_format_multi_slides.pptx", 3);
        var outputPath = CreateTestFilePath("test_format_multi_slides_output.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 2 });
        var result = _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Arial", outputPath: outputPath);
        Assert.Contains("2 slides", result);
    }

    [Fact]
    public void Execute_WithTableText_ShouldFormatTableCells()
    {
        var pptPath = CreatePresentationWithTable("test_format_table.pptx");
        var outputPath = CreateTestFilePath("test_format_table_output.pptx");
        var result = _tool.Execute(pptPath, fontName: "Arial", fontSize: 14, bold: true, outputPath: outputPath);
        Assert.Contains("1 slides", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithMixedShapes_ShouldFormatBothAutoShapeAndTable()
    {
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

    [Fact]
    public void Execute_WithNoFormattingOptions_ShouldStillSucceed()
    {
        var pptPath = CreateTestPresentation("test_format_none.pptx");
        var outputPath = CreateTestFilePath("test_format_none_output.pptx");
        var result = _tool.Execute(pptPath, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithAllSlides_ShouldFormatAllSlides()
    {
        var pptPath = CreateTestPresentation("test_format_all_slides.pptx", 3);
        var outputPath = CreateTestFilePath("test_format_all_slides_output.pptx");
        var result = _tool.Execute(pptPath, fontName: "Arial", outputPath: outputPath);
        Assert.Contains("3 slides", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_index.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 99 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Arial"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_neg_index.pptx");
        var slideIndicesJson = JsonSerializer.Serialize(new[] { -1 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(pptPath, slideIndices: slideIndicesJson, fontName: "Arial"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidColor_ShouldDefaultToBlack()
    {
        var pptPath = CreateTestPresentation("test_invalid_color.pptx");
        var outputPath = CreateTestFilePath("test_invalid_color_output.pptx");
        var result = _tool.Execute(pptPath, outputPath: outputPath, color: "InvalidColorName");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Batch formatted text applied to", result);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndicesJson_ShouldThrowJsonException()
    {
        var pptPath = CreateTestPresentation("test_invalid_json.pptx");
        Assert.ThrowsAny<JsonException>(() =>
            _tool.Execute(pptPath, slideIndices: "not valid json", fontName: "Arial"));
    }

    #endregion

    #region Session

    [Fact]
    public void Execute_WithSessionId_ShouldFormatInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, fontName: "Arial", fontSize: 16, bold: true);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithSessionId_ShouldApplyColorInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_color.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, color: "#FF0000");
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
        Assert.True(ppt.Slides.Count > 0);
    }

    [Fact]
    public void Execute_WithSessionId_MultipleFormats_ShouldApplyAll()
    {
        var pptPath = CreateTestPresentation("test_session_multi.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, fontName: "Verdana", fontSize: 18,
            bold: true, italic: true, color: "#0000FF");
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
    }

    [Fact]
    public void Execute_WithSessionId_AndSlideIndices_ShouldFormatSelectedSlides()
    {
        var pptPath = CreateTestPresentation("test_session_slides.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var slideIndicesJson = JsonSerializer.Serialize(new[] { 0, 2 });
        var result = _tool.Execute(sessionId: sessionId, slideIndices: slideIndicesJson, fontName: "Arial");
        Assert.Contains("2 slides", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute(sessionId: "invalid_session_id", fontName: "Arial"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestFilePath("test_path_format.pptx");
        using (var pres1 = new Presentation())
        {
            var shape1 = pres1.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape1.TextFrame.Text = "PathText";
            pres1.Save(pptPath1, SaveFormat.Pptx);
        }

        var pptPath2 = CreateTestFilePath("test_session_format2.pptx");
        using (var pres2 = new Presentation())
        {
            var shape2 = pres2.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape2.TextFrame.Text = "SessionText";
            pres2.Save(pptPath2, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute(pptPath1, sessionId, fontName: "Arial", bold: true);
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
    }

    #endregion
}