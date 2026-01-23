using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("shown", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("hidden", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("5", result.Message);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [Fact]
    public void Execute_WithDefaults()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("1", result.Message);
    }

    #endregion
}
