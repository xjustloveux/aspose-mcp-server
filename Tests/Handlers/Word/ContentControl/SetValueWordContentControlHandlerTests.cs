using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Handlers.Word.ContentControl;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.ContentControl;

/// <summary>
///     Tests for SetValueWordContentControlHandler.
/// </summary>
public class SetValueWordContentControlHandlerTests : WordHandlerTestBase
{
    private readonly SetValueWordContentControlHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeSetValue()
    {
        Assert.Equal("set_value", _handler.Operation);
    }

    [Fact]
    public void Execute_WithPlainText_ShouldSetTextValue()
    {
        var doc = CreateDocumentWithSdt(SdtType.PlainText, "textTag");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "New Text Value" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("New Text Value", successResult.Message);
        AssertModified(context);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Contains("New Text Value", sdt.GetText());
    }

    [Fact]
    public void Execute_WithCheckBox_ShouldSetCheckedState()
    {
        var doc = CreateDocumentWithSdt(SdtType.Checkbox, "checkTag");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "checkTag" },
            { "value", "true" }
        });

        _handler.Execute(context, parameters);

        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.True(sdt.Checked);
    }

    [Fact]
    public void Execute_WithCheckBox_ShouldUncheck()
    {
        var doc = CreateDocumentWithSdt(SdtType.Checkbox, "checkTag");
        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        sdt.Checked = true;

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "false" }
        });

        _handler.Execute(context, parameters);

        Assert.False(sdt.Checked);
    }

    [Fact]
    public void Execute_WithDropDownList_ShouldSelectItem()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline) { Tag = "colorTag" };
        sdt.ListItems.Add(new SdtListItem("Red", "Red"));
        sdt.ListItems.Add(new SdtListItem("Green", "Green"));
        sdt.ListItems.Add(new SdtListItem("Blue", "Blue"));
        builder.InsertNode(sdt);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "Green" }
        });

        _handler.Execute(context, parameters);

        var updatedSdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Equal("Green", updatedSdt.ListItems.SelectedValue?.Value);
    }

    [Fact]
    public void Execute_WithDropDownList_NonExistentItem_ShouldThrowArgumentException()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline);
        sdt.ListItems.Add(new SdtListItem("Red", "Red"));
        builder.InsertNode(sdt);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "Purple" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found in the list", ex.Message);
    }

    [Fact]
    public void Execute_WithCheckBox_InvalidValue_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdt(SdtType.Checkbox, "checkTag");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "invalid" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid checkbox value", ex.Message);
    }

    [Fact]
    public void Execute_WithLockedContents_ShouldThrowInvalidOperationException()
    {
        var doc = CreateDocumentWithSdt(SdtType.PlainText, "lockedTag");
        var sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        sdt.LockContents = true;

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 },
            { "value", "Should fail" }
        });

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingValue_ShouldThrowArgumentException()
    {
        var doc = CreateDocumentWithSdt(SdtType.PlainText, "tag1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "index", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    private static Document CreateDocumentWithSdt(SdtType sdtType, string tag)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, sdtType, MarkupLevel.Inline) { Tag = tag };
        builder.InsertNode(sdt);
        return doc;
    }
}
