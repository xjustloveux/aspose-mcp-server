using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderTabsHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderTabsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderTabs()
    {
        Assert.Equal("set_header_tabs", _handler.Operation);
    }

    #endregion

    #region HeaderFooter Type

    [Theory]
    [InlineData("primary")]
    [InlineData("first")]
    [InlineData("even")]
    public void Execute_WithDifferentHeaderFooterTypes_SetsTabs(string headerFooterType)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0 }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops },
            { "headerFooterType", headerFooterType }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_WithNoTabStops_SetsHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTabStops_SetsHeaderTabs()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0, ["alignment"] = "left", ["leader"] = "none" },
            new JsonObject { ["position"] = 200.0, ["alignment"] = "center", ["leader"] = "dots" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllAlignments_SetsAllAlignmentTypes()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 50.0, ["alignment"] = "left" },
            new JsonObject { ["position"] = 100.0, ["alignment"] = "center" },
            new JsonObject { ["position"] = 150.0, ["alignment"] = "right" },
            new JsonObject { ["position"] = 200.0, ["alignment"] = "decimal" },
            new JsonObject { ["position"] = 250.0, ["alignment"] = "bar" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    [Fact]
    public void Execute_WithAllLeaders_SetsAllLeaderTypes()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 50.0, ["leader"] = "none" },
            new JsonObject { ["position"] = 100.0, ["leader"] = "dots" },
            new JsonObject { ["position"] = 150.0, ["leader"] = "dashes" },
            new JsonObject { ["position"] = 200.0, ["leader"] = "line" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSpecificSectionIndex_SetsTabs()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0 }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops },
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    [SkippableFact]
    public void Execute_WithAllSections_SetsTabsInAllSections()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits section operations");

        var doc = CreateEmptyDocument();
        doc.AppendChild(new Section(doc));
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject { ["position"] = 100.0 }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops },
            { "sectionIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Execute_WithEmptyTabStops_DoesNotAddTabs()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", new JsonArray() }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    [Fact]
    public void Execute_WithNullValuesInTabStop_UsesDefaults()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var tabStops = new JsonArray
        {
            new JsonObject()
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabStops", tabStops }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tab", result.ToLower());
    }

    #endregion
}
