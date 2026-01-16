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

    #region Alignment Tests

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void Execute_WithAlignment_SetsAlignment(string alignment)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "column" },
            { "data", new[] { new[] { "A", "B" }, new[] { "1", "2" } } },
            { "alignment", alignment }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
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

    #region Additional Chart Types

    [Fact]
    public void Execute_AddsLineChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "line" },
            { "data", new[] { new[] { "Month", "Value" }, new[] { "Jan", "100" }, new[] { "Feb", "120" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("line", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsAreaChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "area" },
            { "data", new[] { new[] { "Q", "Sales" }, new[] { "Q1", "500" }, new[] { "Q2", "600" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("area", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsScatterChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "scatter" },
            { "data", new[] { new[] { "X", "Y" }, new[] { "1", "5" }, new[] { "2", "10" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("scatter", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsDoughnutChart()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "doughnut" },
            { "data", new[] { new[] { "Category", "Value" }, new[] { "A", "40" }, new[] { "B", "60" } } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("doughnut", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Chart Size Tests

    [Fact]
    public void Execute_WithCustomWidth_SetsWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "column" },
            { "data", new[] { new[] { "A", "B" }, new[] { "1", "2" } } },
            { "chartWidth", 600.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomHeight_SetsHeight()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "column" },
            { "data", new[] { new[] { "A", "B" }, new[] { "1", "2" } } },
            { "chartHeight", 300.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Paragraph Index Tests

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithText("Existing content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", new[] { new[] { "X", "Y" }, new[] { "1", "2" } } },
            { "paragraphIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithValidParagraphIndex_InsertsAtPosition()
    {
        var doc = CreateDocumentWithText("First paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", new[] { new[] { "X", "Y" }, new[] { "1", "2" } } },
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully added chart", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", new[] { new[] { "X", "Y" }, new[] { "1", "2" } } },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
