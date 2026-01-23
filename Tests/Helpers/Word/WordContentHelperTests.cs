using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordContentHelperTests
{
    #region CleanText Tests - Null and Empty

    [Fact]
    public void CleanText_WithNull_ReturnsNull()
    {
        var result = WordContentHelper.CleanText(null!);

        Assert.Null(result);
    }

    [Fact]
    public void CleanText_WithEmpty_ReturnsEmpty()
    {
        var result = WordContentHelper.CleanText(string.Empty);

        Assert.Equal(string.Empty, result);
    }

    #endregion

    #region CleanText Tests - Control Characters

    [Fact]
    public void CleanText_WithControlCharacters_RemovesControlCharacters()
    {
        var input = "Hello\u0000\u0001\u0002World";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("HelloWorld", result);
    }

    [Fact]
    public void CleanText_WithCarriageReturn_RemovesCarriageReturn()
    {
        var input = "Hello\r\nWorld";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello\nWorld", result);
    }

    [Fact]
    public void CleanText_PreservesNewlines()
    {
        var input = "Line1\nLine2";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Line1\nLine2", result);
    }

    #endregion

    #region CleanText Tests - Whitespace Normalization

    [Fact]
    public void CleanText_WithMultipleSpaces_NormalizesToSingleSpace()
    {
        var input = "Hello    World";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void CleanText_WithTabs_NormalizesToSpace()
    {
        var input = "Hello\t\tWorld";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void CleanText_WithLeadingAndTrailingWhitespace_Trims()
    {
        var input = "   Hello World   ";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void CleanText_WithSpaceAfterNewline_RemovesSpace()
    {
        var input = "Hello\n   World";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello\nWorld", result);
    }

    #endregion

    #region CleanText Tests - Consecutive Newlines

    [Fact]
    public void CleanText_WithConsecutiveNewlines_PreservesMaxTwo()
    {
        var input = "Hello\n\n\n\nWorld";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello\n\nWorld", result);
    }

    [Fact]
    public void CleanText_WithTwoNewlines_PreservesTwoNewlines()
    {
        var input = "Hello\n\nWorld";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello\n\nWorld", result);
    }

    #endregion

    #region CleanText Tests - Mixed Content

    [Fact]
    public void CleanText_WithMixedContent_CleansCorrectly()
    {
        var input = "  Hello\r\n\r\n  World\u0000\t\tTest  ";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello\n\nWorld Test", result);
    }

    [Fact]
    public void CleanText_WithOnlyWhitespace_ReturnsEmpty()
    {
        var input = "   \t\t   ";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void CleanText_WithNormalText_ReturnsUnchanged()
    {
        var input = "Hello World";

        var result = WordContentHelper.CleanText(input);

        Assert.Equal("Hello World", result);
    }

    #endregion
}
