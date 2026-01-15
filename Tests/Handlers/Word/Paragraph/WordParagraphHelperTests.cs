using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class WordParagraphHelperTests
{
    #region GetAlignment Tests

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("LEFT", ParagraphAlignment.Left)]
    [InlineData("Left", ParagraphAlignment.Left)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("CENTER", ParagraphAlignment.Center)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("RIGHT", ParagraphAlignment.Right)]
    [InlineData("justify", ParagraphAlignment.Justify)]
    [InlineData("JUSTIFY", ParagraphAlignment.Justify)]
    public void GetAlignment_WithValidValues_ReturnsCorrectAlignment(string input, ParagraphAlignment expected)
    {
        var result = WordParagraphHelper.GetAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("centered")]
    public void GetAlignment_WithInvalidValues_ReturnsLeft(string input)
    {
        var result = WordParagraphHelper.GetAlignment(input);

        Assert.Equal(ParagraphAlignment.Left, result);
    }

    #endregion

    #region GetLineSpacingRule Tests

    [Theory]
    [InlineData("atleast", LineSpacingRule.AtLeast)]
    [InlineData("ATLEAST", LineSpacingRule.AtLeast)]
    [InlineData("AtLeast", LineSpacingRule.AtLeast)]
    [InlineData("exactly", LineSpacingRule.Exactly)]
    [InlineData("EXACTLY", LineSpacingRule.Exactly)]
    public void GetLineSpacingRule_WithValidValues_ReturnsCorrectRule(string input, LineSpacingRule expected)
    {
        var result = WordParagraphHelper.GetLineSpacingRule(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("multiple")]
    [InlineData("invalid")]
    [InlineData("")]
    [InlineData("single")]
    public void GetLineSpacingRule_WithInvalidOrDefaultValues_ReturnsMultiple(string input)
    {
        var result = WordParagraphHelper.GetLineSpacingRule(input);

        Assert.Equal(LineSpacingRule.Multiple, result);
    }

    #endregion

    #region GetTabAlignment Tests

    [Theory]
    [InlineData("left", TabAlignment.Left)]
    [InlineData("LEFT", TabAlignment.Left)]
    [InlineData("center", TabAlignment.Center)]
    [InlineData("CENTER", TabAlignment.Center)]
    [InlineData("right", TabAlignment.Right)]
    [InlineData("RIGHT", TabAlignment.Right)]
    [InlineData("decimal", TabAlignment.Decimal)]
    [InlineData("DECIMAL", TabAlignment.Decimal)]
    [InlineData("bar", TabAlignment.Bar)]
    [InlineData("BAR", TabAlignment.Bar)]
    [InlineData("clear", TabAlignment.Clear)]
    [InlineData("CLEAR", TabAlignment.Clear)]
    public void GetTabAlignment_WithValidValues_ReturnsCorrectAlignment(string input, TabAlignment expected)
    {
        var result = WordParagraphHelper.GetTabAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("centered")]
    public void GetTabAlignment_WithInvalidValues_ReturnsLeft(string input)
    {
        var result = WordParagraphHelper.GetTabAlignment(input);

        Assert.Equal(TabAlignment.Left, result);
    }

    #endregion

    #region GetTabLeader Tests

    [Theory]
    [InlineData("none", TabLeader.None)]
    [InlineData("NONE", TabLeader.None)]
    [InlineData("dots", TabLeader.Dots)]
    [InlineData("DOTS", TabLeader.Dots)]
    [InlineData("dashes", TabLeader.Dashes)]
    [InlineData("DASHES", TabLeader.Dashes)]
    [InlineData("line", TabLeader.Line)]
    [InlineData("LINE", TabLeader.Line)]
    [InlineData("heavy", TabLeader.Heavy)]
    [InlineData("HEAVY", TabLeader.Heavy)]
    [InlineData("middledot", TabLeader.MiddleDot)]
    [InlineData("MIDDLEDOT", TabLeader.MiddleDot)]
    public void GetTabLeader_WithValidValues_ReturnsCorrectLeader(string input, TabLeader expected)
    {
        var result = WordParagraphHelper.GetTabLeader(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("spaces")]
    public void GetTabLeader_WithInvalidValues_ReturnsNone(string input)
    {
        var result = WordParagraphHelper.GetTabLeader(input);

        Assert.Equal(TabLeader.None, result);
    }

    #endregion
}
