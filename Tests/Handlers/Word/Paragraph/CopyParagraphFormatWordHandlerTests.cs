using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class CopyParagraphFormatWordHandlerTests : WordHandlerTestBase
{
    private readonly CopyParagraphFormatWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopyFormat()
    {
        Assert.Equal("copy_format", _handler.Operation);
    }

    #endregion

    #region Preserve Content

    [Fact]
    public void Execute_PreservesTargetContent()
    {
        var doc = CreateDocumentWithParagraphs("Source text", "Target text to preserve");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Target text to preserve");
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSourceAndTargetIndices()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
        Assert.Contains("#1", result);
    }

    #endregion

    #region Basic Copy Format Operations

    [Fact]
    public void Execute_CopiesFormat()
    {
        var doc = CreateDocumentWithParagraphs("Source paragraph", "Target paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(0, 2)]
    public void Execute_CopiesBetweenVariousParagraphs(int source, int target)
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", source },
            { "targetParagraphIndex", target }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Format Properties

    [Fact]
    public void Execute_CopiesAlignment()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var sourcePara = paragraphs[0] as Aspose.Words.Paragraph;
        sourcePara!.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        var targetPara = paragraphs[1] as Aspose.Words.Paragraph;
        Assert.Equal(ParagraphAlignment.Center, targetPara!.ParagraphFormat.Alignment);
    }

    [Fact]
    public void Execute_CopiesIndentation()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var sourcePara = paragraphs[0] as Aspose.Words.Paragraph;
        sourcePara!.ParagraphFormat.LeftIndent = 36.0;
        sourcePara.ParagraphFormat.RightIndent = 18.0;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        var targetPara = paragraphs[1] as Aspose.Words.Paragraph;
        Assert.Equal(36.0, targetPara!.ParagraphFormat.LeftIndent);
        Assert.Equal(18.0, targetPara.ParagraphFormat.RightIndent);
    }

    [Fact]
    public void Execute_CopiesSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        var sourcePara = paragraphs[0] as Aspose.Words.Paragraph;
        sourcePara!.ParagraphFormat.SpaceBefore = 12.0;
        sourcePara.ParagraphFormat.SpaceAfter = 6.0;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        var targetPara = paragraphs[1] as Aspose.Words.Paragraph;
        Assert.Equal(12.0, targetPara!.ParagraphFormat.SpaceBefore);
        Assert.Equal(6.0, targetPara.ParagraphFormat.SpaceAfter);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSourceParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutTargetParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSourceIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 100 },
            { "targetParagraphIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Source", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTargetIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", 0 },
            { "targetParagraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Target", ex.Message);
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    public void Execute_WithNegativeIndex_ThrowsArgumentException(int source, int target)
    {
        var doc = CreateDocumentWithParagraphs("Source", "Target");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceParagraphIndex", source },
            { "targetParagraphIndex", target }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
