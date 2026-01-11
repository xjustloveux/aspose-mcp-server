using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddChartWordHandlerTests : WordHandlerTestBase
{
    private readonly AddChartWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddChart()
    {
        Assert.Equal("add_chart", _handler.Operation);
    }

    #endregion

    #region Basic Add Chart Operations

    [Fact]
    public void Execute_AddsColumnChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "column" },
            { "data", new[] { new[] { "Category", "Value" }, new[] { "A", "10" }, new[] { "B", "20" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsBarChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "bar" },
            { "data", new[] { new[] { "Cat", "Val" }, new[] { "X", "15" }, new[] { "Y", "25" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("bar", result.ToLower());
    }

    [Fact]
    public void Execute_AddsPieChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "pie" },
            { "data", new[] { new[] { "Item", "Value" }, new[] { "A", "30" }, new[] { "B", "70" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("pie", result.ToLower());
    }

    [Fact]
    public void Execute_WithChartTitle()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", new[] { new[] { "X", "Y" }, new[] { "1", "2" } } },
            { "chartTitle", "My Chart" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
    }

    [Fact]
    public void Execute_WithoutData_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyData_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", Array.Empty<string[]>() }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
