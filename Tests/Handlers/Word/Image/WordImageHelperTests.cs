using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class WordImageHelperTests : WordTestBase
{
    #region GetAlignment Tests

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("LEFT", ParagraphAlignment.Left)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("CENTER", ParagraphAlignment.Center)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("RIGHT", ParagraphAlignment.Right)]
    public void GetAlignment_WithValidValues_ReturnsCorrectAlignment(string input, ParagraphAlignment expected)
    {
        var result = WordImageHelper.GetAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("justify")]
    [InlineData("")]
    public void GetAlignment_WithInvalidValues_ReturnsLeft(string input)
    {
        var result = WordImageHelper.GetAlignment(input);

        Assert.Equal(ParagraphAlignment.Left, result);
    }

    #endregion

    #region GetWrapType Tests

    [Theory]
    [InlineData("inline", WrapType.Inline)]
    [InlineData("INLINE", WrapType.Inline)]
    [InlineData("square", WrapType.Square)]
    [InlineData("SQUARE", WrapType.Square)]
    [InlineData("tight", WrapType.Tight)]
    [InlineData("through", WrapType.Through)]
    [InlineData("topandbottom", WrapType.TopBottom)]
    [InlineData("none", WrapType.None)]
    public void GetWrapType_WithValidValues_ReturnsCorrectWrapType(string input, WrapType expected)
    {
        var result = WordImageHelper.GetWrapType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("")]
    [InlineData("around")]
    public void GetWrapType_WithInvalidValues_ReturnsInline(string input)
    {
        var result = WordImageHelper.GetWrapType(input);

        Assert.Equal(WrapType.Inline, result);
    }

    #endregion

    #region GetHorizontalAlignment Tests

    [Theory]
    [InlineData("left", HorizontalAlignment.Left)]
    [InlineData("LEFT", HorizontalAlignment.Left)]
    [InlineData("center", HorizontalAlignment.Center)]
    [InlineData("CENTER", HorizontalAlignment.Center)]
    [InlineData("right", HorizontalAlignment.Right)]
    [InlineData("RIGHT", HorizontalAlignment.Right)]
    public void GetHorizontalAlignment_WithValidValues_ReturnsCorrectAlignment(string input,
        HorizontalAlignment expected)
    {
        var result = WordImageHelper.GetHorizontalAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("")]
    public void GetHorizontalAlignment_WithInvalidValues_ReturnsLeft(string input)
    {
        var result = WordImageHelper.GetHorizontalAlignment(input);

        Assert.Equal(HorizontalAlignment.Left, result);
    }

    #endregion

    #region GetVerticalAlignment Tests

    [Theory]
    [InlineData("top", VerticalAlignment.Top)]
    [InlineData("TOP", VerticalAlignment.Top)]
    [InlineData("center", VerticalAlignment.Center)]
    [InlineData("CENTER", VerticalAlignment.Center)]
    [InlineData("bottom", VerticalAlignment.Bottom)]
    [InlineData("BOTTOM", VerticalAlignment.Bottom)]
    public void GetVerticalAlignment_WithValidValues_ReturnsCorrectAlignment(string input, VerticalAlignment expected)
    {
        var result = WordImageHelper.GetVerticalAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("")]
    [InlineData("middle")]
    public void GetVerticalAlignment_WithInvalidValues_ReturnsTop(string input)
    {
        var result = WordImageHelper.GetVerticalAlignment(input);

        Assert.Equal(VerticalAlignment.Top, result);
    }

    #endregion

    #region GetAllImages Tests

    [Fact]
    public void GetAllImages_WithNoImages_ReturnsEmptyList()
    {
        var doc = new Document();

        var result = WordImageHelper.GetAllImages(doc, -1);

        Assert.Empty(result);
    }

    [Fact]
    public void GetAllImages_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = new Document();

        var ex = Assert.Throws<ArgumentException>(() =>
            WordImageHelper.GetAllImages(doc, 10));

        Assert.Contains("Section index", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetAllImages_WithValidSectionIndex_ReturnsImagesFromSection()
    {
        var doc = new Document();

        var result = WordImageHelper.GetAllImages(doc, 0);

        Assert.Empty(result);
    }

    #endregion

    #region InsertCaption Tests

    [Fact]
    public void InsertCaption_WithCenterAlignment_InsertsCaption()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordImageHelper.InsertCaption(builder, "Test Caption", "center");

        var text = doc.GetText();
        Assert.Contains("Figure", text);
        Assert.Contains("Test Caption", text);
    }

    [Fact]
    public void InsertCaption_WithLeftAlignment_InsertsCaption()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordImageHelper.InsertCaption(builder, "Left Caption", "left");

        var text = doc.GetText();
        Assert.Contains("Left Caption", text);
    }

    [Fact]
    public void InsertCaption_WithRightAlignment_InsertsCaption()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordImageHelper.InsertCaption(builder, "Right Caption", "right");

        var text = doc.GetText();
        Assert.Contains("Right Caption", text);
    }

    #endregion
}
