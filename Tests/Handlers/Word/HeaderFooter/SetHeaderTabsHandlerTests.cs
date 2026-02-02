using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var hfType = headerFooterType.ToLower() switch
            {
                "first" => HeaderFooterType.HeaderFirst,
                "even" => HeaderFooterType.HeaderEven,
                _ => HeaderFooterType.HeaderPrimary
            };
            var header = doc.FirstSection.HeadersFooters[hfType];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(100.0, para.ParagraphFormat.TabStops[0].Position);
        }
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_WithNoTabStops_SetsHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(2, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(100.0, para.ParagraphFormat.TabStops[0].Position);
            Assert.Equal(TabAlignment.Left, para.ParagraphFormat.TabStops[0].Alignment);
            Assert.Equal(TabLeader.None, para.ParagraphFormat.TabStops[0].Leader);
            Assert.Equal(200.0, para.ParagraphFormat.TabStops[1].Position);
            Assert.Equal(TabAlignment.Center, para.ParagraphFormat.TabStops[1].Alignment);
            Assert.Equal(TabLeader.Dots, para.ParagraphFormat.TabStops[1].Leader);
        }

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(5, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(TabAlignment.Left, para.ParagraphFormat.TabStops[0].Alignment);
            Assert.Equal(TabAlignment.Center, para.ParagraphFormat.TabStops[1].Alignment);
            Assert.Equal(TabAlignment.Right, para.ParagraphFormat.TabStops[2].Alignment);
            Assert.Equal(TabAlignment.Decimal, para.ParagraphFormat.TabStops[3].Alignment);
            Assert.Equal(TabAlignment.Bar, para.ParagraphFormat.TabStops[4].Alignment);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(4, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(TabLeader.None, para.ParagraphFormat.TabStops[0].Leader);
            Assert.Equal(TabLeader.Dots, para.ParagraphFormat.TabStops[1].Leader);
            Assert.Equal(TabLeader.Dashes, para.ParagraphFormat.TabStops[2].Leader);
            Assert.Equal(TabLeader.Line, para.ParagraphFormat.TabStops[3].Leader);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(100.0, para.ParagraphFormat.TabStops[0].Position);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        foreach (var section in doc.Sections.Cast<Section>())
        {
            var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(100.0, para.ParagraphFormat.TabStops[0].Position);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            Assert.NotNull(header);
            var para = header.FirstParagraph;
            Assert.NotNull(para);
            Assert.Equal(1, para.ParagraphFormat.TabStops.Count);
            Assert.Equal(0.0, para.ParagraphFormat.TabStops[0].Position);
            Assert.Equal(TabAlignment.Left, para.ParagraphFormat.TabStops[0].Alignment);
            Assert.Equal(TabLeader.None, para.ParagraphFormat.TabStops[0].Leader);
        }
    }

    #endregion
}
