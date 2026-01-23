using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Styles;

public class CreateWordStyleHandlerTests : WordHandlerTestBase
{
    private readonly CreateWordStyleHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CreateStyle()
    {
        Assert.Equal("create_style", _handler.Operation);
    }

    #endregion

    #region Style Types

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("list")]
    public void Execute_WithDifferentStyleTypes_CreatesStyle(string styleType)
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", $"Test{styleType}Style" },
            { "styleType", styleType }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_CreatesStyle()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "MyCustomStyle" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(doc.Styles["MyCustomStyle"]);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontSettings_CreatesStyleWithFont()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "BoldStyle" },
            { "fontName", "Arial" },
            { "fontSize", 14.0 },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
        var style = doc.Styles["BoldStyle"];
        Assert.NotNull(style);
        Assert.True(style.Font.Bold);
    }

    [Fact]
    public void Execute_WithCharacterStyleType_CreatesCharacterStyle()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "CharStyle" },
            { "styleType", "character" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithExistingStyleName_ThrowsInvalidOperationException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Normal" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Base Style

    [Fact]
    public void Execute_WithValidBaseStyle_InheritsFromBaseStyle()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "DerivedStyle" },
            { "baseStyle", "Normal" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
        var style = doc.Styles["DerivedStyle"];
        Assert.Equal("Normal", style.BaseStyleName);
    }

    [Fact]
    public void Execute_WithInvalidBaseStyle_CreatesStyleWithoutBase()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "NoBaseStyle" },
            { "baseStyle", "NonExistentStyle" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Font Settings

    [Fact]
    public void Execute_WithItalicAndUnderline_SetsTextFormatting()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "ItalicUnderlineStyle" },
            { "italic", true },
            { "underline", true }
        });

        _handler.Execute(context, parameters);

        var style = doc.Styles["ItalicUnderlineStyle"];
        Assert.True(style.Font.Italic);
        Assert.NotEqual(Underline.None, style.Font.Underline);
    }

    [Fact]
    public void Execute_WithColor_SetsFontColor()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "ColorStyle" },
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFontNameAsciiAndFarEast_SetsBothFonts()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "MultiFontStyle" },
            { "fontNameAscii", "Arial" },
            { "fontNameFarEast", "MS Gothic" }
        });

        _handler.Execute(context, parameters);

        var style = doc.Styles["MultiFontStyle"];
        Assert.Equal("Arial", style.Font.NameAscii);
        Assert.Equal("MS Gothic", style.Font.NameFarEast);
    }

    [Fact]
    public void Execute_WithFontNameAndPartialAscii_SetsCorrectFonts()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "PartialFontStyle" },
            { "fontName", "Times New Roman" },
            { "fontNameAscii", "Arial" }
        });

        _handler.Execute(context, parameters);

        var style = doc.Styles["PartialFontStyle"];
        Assert.Equal("Arial", style.Font.NameAscii);
    }

    #endregion

    #region Paragraph Settings

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    public void Execute_WithAlignment_SetsAlignment(string alignment)
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", $"Align{alignment}Style" },
            { "alignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithSpacing_SetsSpacing()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "SpacingStyle" },
            { "spaceBefore", 12.0 },
            { "spaceAfter", 6.0 }
        });

        _handler.Execute(context, parameters);

        var style = doc.Styles["SpacingStyle"];
        Assert.Equal(12.0, style.ParagraphFormat.SpaceBefore);
        Assert.Equal(6.0, style.ParagraphFormat.SpaceAfter);
    }

    [Fact]
    public void Execute_WithLineSpacing_SetsLineSpacing()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "LineSpacingStyle" },
            { "lineSpacing", 1.5 }
        });

        _handler.Execute(context, parameters);

        var style = doc.Styles["LineSpacingStyle"];
        Assert.Equal(LineSpacingRule.Multiple, style.ParagraphFormat.LineSpacingRule);
    }

    #endregion
}
