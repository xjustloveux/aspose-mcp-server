using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.Equal(InchToPoint(0.5 * level), paragraphs[0].ParagraphFormat.LeftIndent);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.Equal(indent, paragraphs[0].ParagraphFormat.LeftIndent);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.Equal(indent, paragraphs[0].ParagraphFormat.FirstLineIndent);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        Assert.Equal(36.0, paragraphs[0].ParagraphFormat.LeftIndent);
        Assert.Equal(18.0, paragraphs[0].ParagraphFormat.FirstLineIndent);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("No change parameters provided", result.Message);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[0];
        Assert.True(para.ListFormat.IsListItem);
        var level = para.ListFormat.ListLevelNumber;
        Assert.Equal(NumberStyle.UppercaseRoman, para.ListFormat.List.ListLevels[level].NumberStyle);
        AssertModified(context);
    }

    [Theory]
    [InlineData("arabic", NumberStyle.Arabic)]
    [InlineData("roman", NumberStyle.UppercaseRoman)]
    [InlineData("letter", NumberStyle.UppercaseLetter)]
    [InlineData("bullet", NumberStyle.Bullet)]
    public void Execute_WithVariousNumberStyles_ReturnsStyleInfo(string style, NumberStyle expected)
    {
        var doc = CreateDocumentWithList();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "numberStyle", style }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var para = paragraphs[0];
        Assert.True(para.ListFormat.IsListItem);
        var level = para.ListFormat.ListLevelNumber;
        Assert.Equal(expected, para.ListFormat.List.ListLevels[level].NumberStyle);
        AssertModified(context);
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
