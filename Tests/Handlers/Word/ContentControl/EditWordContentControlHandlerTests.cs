using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Handlers.Word.ContentControl;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.ContentControl;

/// <summary>
///     Tests for EditWordContentControlHandler.
/// </summary>
public class EditWordContentControlHandlerTests : WordHandlerTestBase
{
    private readonly EditWordContentControlHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeEdit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    [Fact]
    public void Execute_WithNewTag_ShouldUpdateTag()
    {
        var doc = CreateDocumentWithSdt("oldTag", "Old Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "newTag", "newTag" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("newTag", successResult.Message);
        AssertModified(context);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Equal("newTag", sdt.Tag);
    }

    [Fact]
    public void Execute_WithNewTitle_ShouldUpdateTitle()
    {
        var doc = CreateDocumentWithSdt("tag1", "Old Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "newTitle", "New Title" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Equal("New Title", sdt.Title);
    }

    [Fact]
    public void Execute_WithLockContents_ShouldSetLock()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "lockContents", true }
        });

        _handler.Execute(context, parameters);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.True(sdt.LockContents);
    }

    [Fact]
    public void Execute_WithLockDeletion_ShouldSetLock()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "lockDeletion", true }
        });

        _handler.Execute(context, parameters);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.True(sdt.LockContentControl);
    }

    [Fact]
    public void Execute_ByTag_ShouldFindAndUpdate()
    {
        var doc = CreateDocumentWithSdt("myTag", "My Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "myTag" },
            { "newTitle", "Updated Title" }
        });

        _handler.Execute(context, parameters);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Equal("Updated Title", sdt.Title);
    }

    [Fact]
    public void Execute_WithNoChanges_ShouldReturnNoChangesMessage()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("No changes", successResult.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 99 },
            { "newTag", "newTag" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentTag_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "nonexistent" },
            { "newTag", "newTag" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoIdentifier_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdt("tag1", "Title");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "newTag", "newTag" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    private static Document CreateDocumentWithSdt(string tag, string title)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Tag = tag,
            Title = title
        };
        builder.InsertNode(sdt);
        return doc;
    }
}
