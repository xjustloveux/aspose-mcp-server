using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Handlers.Word.ContentControl;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.ContentControl;

/// <summary>
///     Tests for AddWordContentControlHandler.
/// </summary>
public class AddWordContentControlHandlerTests : WordHandlerTestBase
{
    private readonly AddWordContentControlHandler _handler = new();

    [Fact]
    public void Operation_ShouldBeAdd()
    {
        Assert.Equal("add", _handler.Operation);
    }

    [Fact]
    public void Execute_WithPlainTextType_ShouldAddContentControl()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "PlainText" },
            { "tag", "testTag" },
            { "title", "Test Title" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("PlainText", successResult.Message);
        Assert.Contains("testTag", successResult.Message);
        AssertModified(context);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.True(sdtNodes.Count > 0);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Equal("testTag", sdt.Tag);
        Assert.Equal("Test Title", sdt.Title);
        Assert.Equal(SdtType.PlainText, sdt.SdtType);
    }

    [Fact]
    public void Execute_WithCheckBoxType_ShouldAddCheckBox()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "CheckBox" },
            { "tag", "agreeBox" },
            { "value", "true" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Equal(SdtType.Checkbox, sdt.SdtType);
        Assert.True(sdt.Checked);
    }

    [Fact]
    public void Execute_WithDropDownListType_ShouldAddDropDownWithItems()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "DropDownList" },
            { "tag", "colorPicker" },
            { "items", "Red,Green,Blue" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Equal(SdtType.DropDownList, sdt.SdtType);
        Assert.Equal(3, sdt.ListItems.Count);
    }

    [Fact]
    public void Execute_WithLockOptions_ShouldSetLocks()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "PlainText" },
            { "lockContents", true },
            { "lockDeletion", true }
        });

        _handler.Execute(context, parameters);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.True(sdt.LockContents);
        Assert.True(sdt.LockContentControl);
    }

    [Fact]
    public void Execute_WithPlainTextAndValue_ShouldSetText()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "PlainText" },
            { "value", "Hello World" }
        });

        _handler.Execute(context, parameters);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Contains("Hello World", sdt.GetText());
    }

    [Fact]
    public void Execute_WithMissingType_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tag", "noType" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnknownType_ShouldThrowArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", "InvalidType" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unknown content control type", ex.Message);
    }

    [Theory]
    [InlineData("DatePicker", SdtType.Date)]
    [InlineData("ComboBox", SdtType.ComboBox)]
    [InlineData("Picture", SdtType.Picture)]
    [InlineData("RichText", SdtType.RichText)]
    public void Execute_WithVariousTypes_ShouldCreateCorrectType(string typeStr, SdtType expectedType)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "type", typeStr }
        });

        _handler.Execute(context, parameters);

        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Equal(expectedType, sdt.SdtType);
    }
}
