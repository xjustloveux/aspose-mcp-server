using Aspose.Words;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Tests.Helpers.Word;

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

    /// <summary>
    ///     A value this helper does not know is refused, not quietly turned into the default.
    ///     <para>
    ///         It used to return Left, so <c>alignment: "centered"</c> left-aligned the paragraph
    ///         and reported success — the caller was never told their value meant nothing
    ///         (§23.13.1).
    ///     </para>
    /// </summary>
    /// <param name="input">A value outside the documented set.</param>
    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("centered")]
    public void GetAlignment_WithUnknownValues_IsRefused(string input)
    {
        var refusal = Assert.Throws<ArgumentException>(() => WordParagraphHelper.GetAlignment(input));

        Assert.Contains("Unknown alignment", refusal.Message, StringComparison.Ordinal);
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

    /// <summary>
    ///     The four multiple-based rules the tool documents all map to Multiple; anything else is
    ///     refused. The whitelist comes from <c>word_paragraph</c>'s own parameter description —
    ///     writing it from the switch arms instead is how <c>double</c>, a value in real use,
    ///     nearly became a refusal (§23.13.1).
    /// </summary>
    /// <param name="input">A documented multiple-based rule.</param>
    [Theory]
    [InlineData("multiple")]
    [InlineData("single")]
    [InlineData("oneAndHalf")]
    [InlineData("double")]
    public void GetLineSpacingRule_WithDocumentedMultipleRules_ReturnsMultiple(string input)
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

    /// <summary>A tab alignment outside the documented set is refused.</summary>
    /// <param name="input">A value outside the documented set.</param>
    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("centered")]
    public void GetTabAlignment_WithUnknownValues_IsRefused(string input)
    {
        var refusal = Assert.Throws<ArgumentException>(() => WordParagraphHelper.GetTabAlignment(input));

        Assert.Contains("Unknown tab alignment", refusal.Message, StringComparison.Ordinal);
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

    /// <summary>A tab leader outside the documented set is refused.</summary>
    /// <param name="input">A value outside the documented set.</param>
    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("spaces")]
    public void GetTabLeader_WithUnknownValues_IsRefused(string input)
    {
        var refusal = Assert.Throws<ArgumentException>(() => WordParagraphHelper.GetTabLeader(input));

        Assert.Contains("Unknown tab leader", refusal.Message, StringComparison.Ordinal);
    }

    #endregion
}
