using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordFormatHelperTests : WordTestBase
{
    #region Helper Methods

    private static Document CreateDocumentWithParagraphs(params string[] texts)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 0; i < texts.Length; i++)
        {
            builder.Write(texts[i]);
            if (i < texts.Length - 1)
                builder.InsertParagraph();
        }

        return doc;
    }

    #endregion

    #region GetTargetParagraph Tests - Valid Index

    [Fact]
    public void GetTargetParagraph_WithIndexZero_ReturnsFirstParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");

        var result = WordFormatHelper.GetTargetParagraph(doc, 0);

        Assert.Contains("First", result.GetText());
    }

    [Fact]
    public void GetTargetParagraph_WithPositiveIndex_ReturnsCorrectParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");

        var result = WordFormatHelper.GetTargetParagraph(doc, 1);

        Assert.Contains("Second", result.GetText());
    }

    [Fact]
    public void GetTargetParagraph_WithLastIndex_ReturnsLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");

        var result = WordFormatHelper.GetTargetParagraph(doc, 2);

        Assert.Contains("Third", result.GetText());
    }

    [Fact]
    public void GetTargetParagraph_WithMinusOne_ReturnsLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");

        var result = WordFormatHelper.GetTargetParagraph(doc, -1);

        Assert.Contains("Third", result.GetText());
    }

    #endregion

    #region GetTargetParagraph Tests - Invalid Cases

    [Fact]
    public void GetTargetParagraph_WithEmptyDocument_ThrowsArgumentException()
    {
        var doc = new Document();
        doc.RemoveAllChildren();

        var ex = Assert.Throws<ArgumentException>(() =>
            WordFormatHelper.GetTargetParagraph(doc, 0));

        Assert.Contains("has no paragraphs", ex.Message);
    }

    [Fact]
    public void GetTargetParagraph_WithIndexTooLarge_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");

        var ex = Assert.Throws<ArgumentException>(() =>
            WordFormatHelper.GetTargetParagraph(doc, 5));

        Assert.Contains("paragraphIndex must be between", ex.Message);
    }

    [Fact]
    public void GetTargetParagraph_WithNegativeIndexOtherThanMinusOne_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");

        var ex = Assert.Throws<ArgumentException>(() =>
            WordFormatHelper.GetTargetParagraph(doc, -2));

        Assert.Contains("paragraphIndex must be between", ex.Message);
    }

    #endregion

    #region GetLineStyle Tests

    [Theory]
    [InlineData("none", LineStyle.None)]
    [InlineData("NONE", LineStyle.None)]
    [InlineData("single", LineStyle.Single)]
    [InlineData("SINGLE", LineStyle.Single)]
    [InlineData("double", LineStyle.Double)]
    [InlineData("DOUBLE", LineStyle.Double)]
    [InlineData("dotted", LineStyle.Dot)]
    [InlineData("DOTTED", LineStyle.Dot)]
    [InlineData("dashed", LineStyle.Single)]
    [InlineData("thick", LineStyle.Thick)]
    [InlineData("THICK", LineStyle.Thick)]
    public void GetLineStyle_WithValidValues_ReturnsCorrectStyle(string input, LineStyle expected)
    {
        var result = WordFormatHelper.GetLineStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void GetLineStyle_WithInvalidValues_ReturnsSingle(string input)
    {
        var result = WordFormatHelper.GetLineStyle(input);

        Assert.Equal(LineStyle.Single, result);
    }

    #endregion

    #region GetColorName Tests

    [Fact]
    public void GetColorName_WithEmptyColor_ReturnsAutoBlack()
    {
        var result = WordFormatHelper.GetColorName(Color.Empty);

        Assert.Equal("Auto/Black", result);
    }

    [Fact]
    public void GetColorName_WithTransparentBlack_ReturnsAutoBlack()
    {
        var color = Color.FromArgb(0, 0, 0, 0);

        var result = WordFormatHelper.GetColorName(color);

        Assert.Equal("Auto/Black", result);
    }

    [Theory]
    [InlineData(255, 0, 0, "Red")]
    [InlineData(0, 255, 0, "Green")]
    [InlineData(0, 0, 255, "Blue")]
    [InlineData(255, 255, 0, "Yellow")]
    [InlineData(255, 0, 255, "Magenta")]
    [InlineData(0, 255, 255, "Cyan")]
    [InlineData(255, 255, 255, "White")]
    [InlineData(128, 128, 128, "Gray")]
    [InlineData(255, 165, 0, "Orange")]
    [InlineData(128, 0, 128, "Purple")]
    public void GetColorName_WithKnownColors_ReturnsCorrectName(int r, int g, int b, string expectedName)
    {
        var color = Color.FromArgb(r, g, b);

        var result = WordFormatHelper.GetColorName(color);

        Assert.Equal(expectedName, result);
    }

    [Fact]
    public void GetColorName_WithUnknownColor_ReturnsCustom()
    {
        var color = Color.FromArgb(123, 45, 67);

        var result = WordFormatHelper.GetColorName(color);

        Assert.Equal("Custom", result);
    }

    #endregion
}
