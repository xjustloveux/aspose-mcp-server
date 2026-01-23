using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Unit tests for FontHelper class and its nested helper classes
/// </summary>
public class FontHelperTests : TestBase
{
    #region Word Font Helper Tests

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithFontName_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, "Arial");

        Assert.Equal("Arial", run.Font.Name);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithFontSize_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, fontSize: 14);

        Assert.Equal(14, run.Font.Size);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithBold_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, bold: true);

        Assert.True(run.Font.Bold);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithItalic_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, italic: true);

        Assert.True(run.Font.Italic);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithUnderline_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, underline: "single");

        Assert.Equal(Underline.Single, run.Font.Underline);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithColor_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, color: "#FF0000");

        Assert.Equal(255, run.Font.Color.R);
        Assert.Equal(0, run.Font.Color.G);
        Assert.Equal(0, run.Font.Color.B);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithStrikethrough_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, strikethrough: true);

        Assert.True(run.Font.StrikeThrough);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithSuperscript_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, superscript: true);

        Assert.True(run.Font.Superscript);
        Assert.False(run.Font.Subscript);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithSubscript_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, subscript: true);

        Assert.True(run.Font.Subscript);
        Assert.False(run.Font.Superscript);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithFontNameAscii_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, fontNameAscii: "Courier New");

        Assert.Equal("Courier New", run.Font.NameAscii);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToRun_WithFontNameFarEast_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var run = (Run)doc.GetChild(NodeType.Run, 0, true);

        FontHelper.Word.ApplyFontSettings(run, fontNameFarEast: "MS Mincho");

        Assert.Equal("MS Mincho", run.Font.NameFarEast);
    }

    [Fact]
    public void Word_ApplyFontSettings_ToBuilder_WithAllSettings_ShouldApply()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        FontHelper.Word.ApplyFontSettings(builder,
            "Arial",
            fontSize: 16,
            bold: true,
            italic: true,
            underline: "double",
            color: "#00FF00",
            strikethrough: true);

        Assert.Equal("Arial", builder.Font.Name);
        Assert.Equal(16, builder.Font.Size);
        Assert.True(builder.Font.Bold);
        Assert.True(builder.Font.Italic);
        Assert.Equal(Underline.Double, builder.Font.Underline);
        Assert.True(builder.Font.StrikeThrough);
    }

    [Fact]
    public void Word_ParseUnderline_WithNull_ShouldReturnNone()
    {
        var result = FontHelper.Word.ParseUnderline(null);

        Assert.Equal(Underline.None, result);
    }

    [Fact]
    public void Word_ParseUnderline_WithEmpty_ShouldReturnNone()
    {
        var result = FontHelper.Word.ParseUnderline("");

        Assert.Equal(Underline.None, result);
    }

    [Theory]
    [InlineData("single", Underline.Single)]
    [InlineData("double", Underline.Double)]
    [InlineData("dotted", Underline.Dotted)]
    [InlineData("dash", Underline.Dash)]
    [InlineData("none", Underline.None)]
    [InlineData("SINGLE", Underline.Single)]
    [InlineData("invalid", Underline.None)]
    public void Word_ParseUnderline_ShouldReturnCorrectValue(string input, Underline expected)
    {
        var result = FontHelper.Word.ParseUnderline(input);

        Assert.Equal(expected, result);
    }

    #endregion

    #region Excel Font Helper Tests

    [Fact]
    public void Excel_ApplyFontSettings_WithFontName_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style, "Calibri");

        Assert.Equal("Calibri", style.Font.Name);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithFontSize_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style, fontSize: 14);

        Assert.Equal(14, style.Font.Size);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithBold_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style, bold: true);

        Assert.True(style.Font.IsBold);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithItalic_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style, italic: true);

        Assert.True(style.Font.IsItalic);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithFontColor_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style, fontColor: "#0000FF");

        Assert.Equal(0, style.Font.Color.R);
        Assert.Equal(0, style.Font.Color.G);
        Assert.Equal(255, style.Font.Color.B);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithAllSettings_ShouldApply()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();

        FontHelper.Excel.ApplyFontSettings(style,
            "Arial",
            12,
            true,
            true,
            "#FF0000");

        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal(12, style.Font.Size);
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
    }

    [Fact]
    public void Excel_ApplyFontSettings_WithNullValues_ShouldNotChange()
    {
        using var workbook = new Workbook();
        var style = workbook.CreateStyle();
        var originalName = style.Font.Name;

        FontHelper.Excel.ApplyFontSettings(style);

        Assert.Equal(originalName, style.Font.Name);
    }

    #endregion

    #region PowerPoint Font Helper Tests

    [Fact]
    public void Ppt_ApplyFontSettings_WithFontName_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, "Arial");

        Assert.Equal("Arial", portion.PortionFormat.LatinFont.FontName);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithFontSize_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, fontSize: 24);

        Assert.Equal(24, portion.PortionFormat.FontHeight);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithBold_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, bold: true);

        Assert.Equal(NullableBool.True, portion.PortionFormat.FontBold);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithItalic_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, italic: true);

        Assert.Equal(NullableBool.True, portion.PortionFormat.FontItalic);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithColor_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, color: "#FF0000");

        Assert.Equal(FillType.Solid, portion.PortionFormat.FillFormat.FillType);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithAllSettings_ShouldApply()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat,
            "Verdana",
            18,
            true,
            true,
            "#00FF00");

        Assert.Equal("Verdana", portion.PortionFormat.LatinFont.FontName);
        Assert.Equal(18, portion.PortionFormat.FontHeight);
        Assert.Equal(NullableBool.True, portion.PortionFormat.FontBold);
        Assert.Equal(NullableBool.True, portion.PortionFormat.FontItalic);
    }

    [Fact]
    public void Ppt_ApplyFontSettings_WithNullValues_ShouldNotChange()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Test";
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];
        var originalFont = portion.PortionFormat.LatinFont;

        FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat);

        // Font should not be changed
        Assert.Equal(originalFont, portion.PortionFormat.LatinFont);
    }

    #endregion

    #region PDF Font Helper Tests

    [Fact]
    public void Pdf_ApplyFontSettings_WithFontSize_ShouldApply()
    {
        var textState = new TextState();

        FontHelper.Pdf.ApplyFontSettings(textState, fontSize: 14);

        Assert.Equal(14, textState.FontSize);
    }

    [Fact]
    public void Pdf_ApplyFontSettings_WithNullValues_ShouldNotChange()
    {
        var textState = new TextState();
        var originalSize = textState.FontSize;

        FontHelper.Pdf.ApplyFontSettings(textState);

        Assert.Equal(originalSize, textState.FontSize);
    }

    [Fact]
    public void Pdf_ApplyFontSettings_WithInvalidFontName_ShouldNotThrow()
    {
        var textState = new TextState();

        var exception = Record.Exception(() =>
            FontHelper.Pdf.ApplyFontSettings(textState, "NonExistentFont12345"));

        Assert.Null(exception);
    }

    #endregion
}
