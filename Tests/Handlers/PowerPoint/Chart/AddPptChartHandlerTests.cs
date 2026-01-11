using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class AddPptChartHandlerTests : PptHandlerTestBase
{
    private readonly AddPptChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsChart()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Column" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("column", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTitle_AddsChartWithTitle()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Pie" },
            { "title", "My Chart" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("pie", result.ToLower());
        Assert.Contains("added", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomPosition_AddsChartAtPosition()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Line" },
            { "x", 100f },
            { "y", 100f },
            { "width", 400f },
            { "height", 300f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "Column" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutChartType_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "chartType", "Column" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
