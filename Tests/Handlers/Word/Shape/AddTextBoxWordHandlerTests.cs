using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddTextBoxWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTextBoxWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddTextbox()
    {
        Assert.Equal("add_textbox", _handler.Operation);
    }

    #endregion

    #region Basic Add TextBox Operations

    [Fact]
    public void Execute_AddsTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added textbox", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithDimensions_AddsTextBoxWithSize()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test Text" },
            { "textboxWidth", 300.0 },
            { "textboxHeight", 150.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added textbox", result.ToLower());
    }

    [Fact]
    public void Execute_WithPosition_AddsTextBoxAtPosition()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned Text" },
            { "positionX", 200.0 },
            { "positionY", 300.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added textbox", result.ToLower());
    }

    [Fact]
    public void Execute_WithBackgroundColor_AddsStyledTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled Text" },
            { "backgroundColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added textbox", result.ToLower());
    }

    [Fact]
    public void Execute_WithFontSettings_AddsFormattedTextBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted Text" },
            { "fontSize", 14.0 },
            { "bold", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added textbox", result.ToLower());
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
