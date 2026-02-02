using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class AddWordListHandlerTests : WordHandlerTestBase
{
    private readonly AddWordListHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddList()
    {
        Assert.Equal("add_list", _handler.Operation);
    }

    #endregion

    #region Number Format Parameter

    [Theory]
    [InlineData("arabic")]
    [InlineData("roman")]
    [InlineData("letter")]
    public void Execute_WithNumberFormat_ReturnsFormatInfo(string format)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items },
            { "listType", "number" },
            { "numberFormat", format }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listPara = paragraphs.First(p => p.ListFormat.IsListItem);
        Assert.NotNull(listPara.ListFormat.List);
        AssertModified(context);
    }

    #endregion

    #region Items with Level Parameter

    [Fact]
    public void Execute_WithObjectItems_SupportsLevels()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray
        {
            new JsonObject { ["text"] = "Main Item", ["level"] = 0 },
            new JsonObject { ["text"] = "Sub Item", ["level"] = 1 }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listItems = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.Equal(2, listItems.Count);
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsListWithItems()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1", "Item 2", "Item 3" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listItems = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.Equal(3, listItems.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsItemCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1", "Item 2" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listItems = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.Equal(2, listItems.Count);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AddsItemTextToDocument()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds watermark to text");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "First Item", "Second Item" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "First Item");
        AssertContainsText(doc, "Second Item");
    }

    #endregion

    #region List Type Parameter

    [Fact]
    public void Execute_WithBulletType_CreatesBulletList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items },
            { "listType", "bullet" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listPara = paragraphs.First(p => p.ListFormat.IsListItem);
        Assert.NotNull(listPara.ListFormat.List);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNumberType_CreatesNumberList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items },
            { "listType", "number" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listPara = paragraphs.First(p => p.ListFormat.IsListItem);
        Assert.NotNull(listPara.ListFormat.List);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomType_ReturnsCustomInfo()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items },
            { "listType", "custom" },
            { "bulletChar", "★" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listPara = paragraphs.First(p => p.ListFormat.IsListItem);
        Assert.NotNull(listPara.ListFormat.List);
        AssertModified(context);
    }

    #endregion

    #region Continue Previous Parameter

    [Fact]
    public void Execute_WithContinuePrevious_ContinuesExistingList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);

        var items1 = new JsonArray { "Item 1", "Item 2" };
        var params1 = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items1 }
        });
        _handler.Execute(context, params1);

        var items2 = new JsonArray { "Item 3", "Item 4" };
        var params2 = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items2 },
            { "continuePrevious", true }
        });

        var res = _handler.Execute(context, params2);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listItems = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.Equal(4, listItems.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithContinuePreviousNoExistingList_CreatesNewList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray { "Item 1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items },
            { "continuePrevious", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();
        var listItems = paragraphs.Where(p => p.ListFormat.IsListItem).ToList();
        Assert.Single(listItems);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutItems_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyItems_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var items = new JsonArray();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "items", items }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cannot be empty", ex.Message);
    }

    #endregion
}
