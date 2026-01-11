using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created successfully", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created", result.ToLower());
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
}
