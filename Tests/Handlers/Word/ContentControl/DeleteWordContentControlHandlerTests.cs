using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Handlers.Word.ContentControl;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.ContentControl;

/// <summary>
///     Tests for DeleteWordContentControlHandler.
/// </summary>
public class DeleteWordContentControlHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordContentControlHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeDelete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    [Fact]
    public void Execute_WithKeepContent_ShouldRemoveControlButKeepText()
    {
        var doc = CreateDocumentWithSdtAndText("deleteTag", "Keep this text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "keepContent", true }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("deleted", successResult.Message);
        Assert.Contains("Content preserved", successResult.Message);
        AssertModified(context);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.Equal(0, sdtNodes.Count);
        AssertContainsText(doc, "Keep this text");
    }

    [Fact]
    public void Execute_WithoutKeepContent_ShouldRemoveControlAndContent()
    {
        var doc = CreateDocumentWithSdtAndText("deleteTag", "Remove this text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "keepContent", false }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("Content removed", successResult.Message);
        AssertModified(context);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.Equal(0, sdtNodes.Count);
    }

    [Fact]
    public void Execute_ByTag_ShouldDeleteCorrectControl()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var sdt1 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "keep" };
        builder.InsertNode(sdt1);
        var sdt2 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "remove" };
        builder.InsertNode(sdt2);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "remove" },
            { "keepContent", true }
        });

        _handler.Execute(context, parameters);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.Single(sdtNodes);
        Assert.Equal("keep", ((StructuredDocumentTag)sdtNodes[0]).Tag);
    }

    [Fact]
    public void Execute_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdtAndText("tag1", "text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentTag_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdtAndText("tag1", "text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "nonexistent" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_DefaultKeepContent_ShouldBeTrue()
    {
        var doc = CreateDocumentWithSdtAndText("defaultTag", "Default text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "keepContent", true }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Default text");
    }

    private static Document CreateDocumentWithSdtAndText(string tag, string text)
    {
        var doc = new Document();
        var sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Tag = tag
        };
        var para = new Aspose.Words.Paragraph(doc);
        para.AppendChild(new Run(doc, text));
        sdt.AppendChild(para);
        doc.FirstSection.Body.AppendChild(sdt);
        return doc;
    }
}
