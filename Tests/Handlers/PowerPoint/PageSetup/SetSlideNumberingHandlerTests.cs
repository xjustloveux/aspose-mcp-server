using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

public class SetSlideNumberingHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideNumberingHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetSlideNumbering()
    {
        Assert.Equal("set_slide_numbering", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Numbering Operations

    [Fact]
    public void Execute_ShowsSlideNumbers()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showSlideNumber", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shown", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_HidesSlideNumbers()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showSlideNumber", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hidden", result.ToLower());
    }

    [Fact]
    public void Execute_SetsFirstNumber()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "firstNumber", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("5", result);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [Fact]
    public void Execute_WithDefaults()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1", result);
    }

    #endregion
}
