using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordShapeHelperTests : WordTestBase
{
    #region FindAllTextboxes Tests

    [Fact]
    public void FindAllTextboxes_WithNoTextboxes_ReturnsEmptyList()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain text");

        var result = WordShapeHelper.FindAllTextboxes(doc);

        Assert.Empty(result);
    }

    [Fact]
    public void FindAllTextboxes_WithTextbox_ReturnsTextbox()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertShape(ShapeType.TextBox, 100, 50);
        shape.AppendChild(new WordParagraph(doc));

        var result = WordShapeHelper.FindAllTextboxes(doc);

        Assert.Single(result);
    }

    [Fact]
    public void FindAllTextboxes_WithMultipleTextboxes_ReturnsAll()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.TextBox, 100, 50);
        builder.InsertShape(ShapeType.TextBox, 100, 50);
        builder.InsertShape(ShapeType.TextBox, 100, 50);

        var result = WordShapeHelper.FindAllTextboxes(doc);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void FindAllTextboxes_WithNonTextboxShapes_ReturnsOnlyTextboxes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.TextBox, 100, 50);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 100, 50);

        var result = WordShapeHelper.FindAllTextboxes(doc);

        Assert.Single(result);
    }

    #endregion

    #region GetAllShapes Tests

    [Fact]
    public void GetAllShapes_WithNoShapes_ReturnsEmptyList()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain text");

        var result = WordShapeHelper.GetAllShapes(doc);

        Assert.Empty(result);
    }

    [Fact]
    public void GetAllShapes_WithShapes_ReturnsAll()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 100, 50);
        builder.InsertShape(ShapeType.TextBox, 100, 50);

        var result = WordShapeHelper.GetAllShapes(doc);

        Assert.Equal(3, result.Count);
    }

    #endregion

    #region ParseDashStyle Tests

    [Theory]
    [InlineData("dash", DashStyle.Dash)]
    [InlineData("DASH", DashStyle.Dash)]
    [InlineData("dot", DashStyle.Dot)]
    [InlineData("DOT", DashStyle.Dot)]
    [InlineData("dashdot", DashStyle.DashDot)]
    [InlineData("DASHDOT", DashStyle.DashDot)]
    [InlineData("dashdotdot", DashStyle.LongDashDotDot)]
    [InlineData("rounddot", DashStyle.ShortDot)]
    public void ParseDashStyle_WithValidValues_ReturnsCorrectStyle(string input, DashStyle expected)
    {
        var result = WordShapeHelper.ParseDashStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("solid")]
    [InlineData("invalid")]
    [InlineData("")]
    [InlineData("unknown")]
    public void ParseDashStyle_WithInvalidOrDefaultValues_ReturnsSolid(string input)
    {
        var result = WordShapeHelper.ParseDashStyle(input);

        Assert.Equal(DashStyle.Solid, result);
    }

    #endregion

    #region ParseShapeType Tests

    [Theory]
    [InlineData("rectangle", ShapeType.Rectangle)]
    [InlineData("RECTANGLE", ShapeType.Rectangle)]
    [InlineData("ellipse", ShapeType.Ellipse)]
    [InlineData("ELLIPSE", ShapeType.Ellipse)]
    [InlineData("roundrectangle", ShapeType.RoundRectangle)]
    [InlineData("line", ShapeType.Line)]
    public void ParseShapeType_WithValidValues_ReturnsCorrectType(string input, ShapeType expected)
    {
        var result = WordShapeHelper.ParseShapeType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void ParseShapeType_WithInvalidValues_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() => WordShapeHelper.ParseShapeType(input));

        Assert.Contains("Unknown shape type", ex.Message);
    }

    #endregion

    #region ParseAlignment Tests

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("LEFT", ParagraphAlignment.Left)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("CENTER", ParagraphAlignment.Center)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("RIGHT", ParagraphAlignment.Right)]
    public void ParseAlignment_WithValidValues_ReturnsCorrectAlignment(string input, ParagraphAlignment expected)
    {
        var result = WordShapeHelper.ParseAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("justify")]
    [InlineData("invalid")]
    [InlineData("")]
    public void ParseAlignment_WithInvalidOrDefaultValues_ReturnsLeft(string input)
    {
        var result = WordShapeHelper.ParseAlignment(input);

        Assert.Equal(ParagraphAlignment.Left, result);
    }

    #endregion
}
