using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Handlers.Word.ContentControl;
using AsposeMcpServer.Results.Word.ContentControl;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.ContentControl;

/// <summary>
///     Tests for GetWordContentControlsHandler.
/// </summary>
public class GetWordContentControlsHandlerTests : WordHandlerTestBase
{
    private readonly GetWordContentControlsHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeGet()
    {
        Assert.Equal("get", _handler.Operation);
    }

    [Fact]
    public void Execute_WithNoContentControls_ShouldReturnEmptyResult()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);

        var result = _handler.Execute(context, CreateEmptyParameters());

        var getResult = Assert.IsType<GetContentControlsResult>(result);
        Assert.Equal(0, getResult.Count);
        Assert.Empty(getResult.ContentControls);
        Assert.NotNull(getResult.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithContentControls_ShouldReturnAll()
    {
        var doc = CreateDocumentWithContentControls(3);
        var context = CreateContext(doc);

        var result = _handler.Execute(context, CreateEmptyParameters());

        var getResult = Assert.IsType<GetContentControlsResult>(result);
        Assert.Equal(3, getResult.Count);
        Assert.Equal(3, getResult.ContentControls.Count);
        Assert.Null(getResult.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ShouldReturnCorrectProperties()
    {
        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Tag = "testTag",
            Title = "Test Title",
            LockContents = true,
            LockContentControl = false
        };
        builder.InsertNode(sdt);

        var context = CreateContext(doc);

        var result = _handler.Execute(context, CreateEmptyParameters());

        var getResult = Assert.IsType<GetContentControlsResult>(result);
        Assert.Single(getResult.ContentControls);
        var info = getResult.ContentControls[0];
        Assert.Equal(0, info.Index);
        Assert.Equal("testTag", info.Tag);
        Assert.Equal("Test Title", info.Title);
        Assert.Equal("PlainText", info.Type);
        Assert.True(info.LockContents);
        Assert.False(info.LockDeletion);
    }

    [Fact]
    public void Execute_WithTagFilter_ShouldFilterByTag()
    {
        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);

        var sdt1 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "alpha" };
        builder.InsertNode(sdt1);
        var sdt2 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "beta" };
        builder.InsertNode(sdt2);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?> { { "tag", "alpha" } });

        var result = _handler.Execute(context, parameters);

        var getResult = Assert.IsType<GetContentControlsResult>(result);
        Assert.Single(getResult.ContentControls);
        Assert.Equal("alpha", getResult.ContentControls[0].Tag);
    }

    [Fact]
    public void Execute_WithTypeFilter_ShouldFilterByType()
    {
        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);

        var sdt1 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "text1" };
        builder.InsertNode(sdt1);
        var sdt2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline) { Tag = "check1" };
        builder.InsertNode(sdt2);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?> { { "type", "Checkbox" } });

        var result = _handler.Execute(context, parameters);

        var getResult = Assert.IsType<GetContentControlsResult>(result);
        Assert.Single(getResult.ContentControls);
        Assert.Equal("Checkbox", getResult.ContentControls[0].Type);
    }

    private static Document CreateDocumentWithContentControls(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Tag = $"tag{i}",
                Title = $"Title {i}"
            };
            builder.InsertNode(sdt);
        }

        return doc;
    }
}
