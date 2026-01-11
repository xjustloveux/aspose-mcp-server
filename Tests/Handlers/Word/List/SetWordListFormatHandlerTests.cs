using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class SetWordListFormatHandlerTests : WordHandlerTestBase
{
    private readonly SetWordListFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFormat()
    {
        Assert.Equal("set_format", _handler.Operation);
    }

    #endregion

    #region Indent Level Parameter

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithIndentLevel_AppliesIndent(int level)
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "indentLevel", level }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Indent level: {level}", result);
    }

    #endregion

    #region Left Indent Parameter

    [Theory]
    [InlineData(0.0)]
    [InlineData(18.0)]
    [InlineData(36.0)]
    public void Execute_WithLeftIndent_AppliesIndent(double indent)
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "leftIndent", indent }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Left indent: {indent}", result);
    }

    #endregion

    #region First Line Indent Parameter

    [Theory]
    [InlineData(0.0)]
    [InlineData(18.0)]
    [InlineData(-18.0)]
    public void Execute_WithFirstLineIndent_AppliesIndent(double indent)
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "firstLineIndent", indent }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"First line indent: {indent}", result);
    }

    #endregion

    #region Multiple Parameters

    [Fact]
    public void Execute_WithMultipleParameters_AppliesAll()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "leftIndent", 36.0 },
            { "firstLineIndent", 18.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Left indent:", result);
        Assert.Contains("First line indent:", result);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithList()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        var list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        builder.Writeln("List Item 1");
        builder.Writeln("List Item 2");
        builder.ListFormat.RemoveNumbers();

        return doc;
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsListFormat()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("List format set successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsParagraphIndex()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Paragraph index: 0", result);
    }

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangeMessage()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No change parameters provided", result);
    }

    #endregion

    #region Number Style Parameter

    [Fact]
    public void Execute_WithNumberStyle_OnListItem_AppliesStyle()
    {
        var doc = CreateDocumentWithList();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "numberStyle", "roman" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Number style: roman", result);
    }

    [Theory]
    [InlineData("arabic")]
    [InlineData("roman")]
    [InlineData("letter")]
    [InlineData("bullet")]
    public void Execute_WithVariousNumberStyles_ReturnsStyleInfo(string style)
    {
        var doc = CreateDocumentWithList();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "numberStyle", style }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Number style: {style}", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
