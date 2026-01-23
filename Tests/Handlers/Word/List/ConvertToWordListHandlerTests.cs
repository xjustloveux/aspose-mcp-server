using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class ConvertToWordListHandlerTests : WordHandlerTestBase
{
    private readonly ConvertToWordListHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ConvertToList()
    {
        Assert.Equal("convert_to_list", _handler.Operation);
    }

    #endregion

    #region Number Format Parameter

    [Theory]
    [InlineData("arabic")]
    [InlineData("roman")]
    [InlineData("letter")]
    public void Execute_WithNumberFormat_AppliesFormat(string format)
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 },
            { "listType", "number" },
            { "numberFormat", format }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"Number format: {format}", result.Message);
    }

    #endregion

    #region Various Ranges

    [Theory]
    [InlineData(0, 0)]
    [InlineData(0, 1)]
    [InlineData(1, 2)]
    [InlineData(0, 2)]
    public void Execute_ConvertsVariousRanges(int start, int end)
    {
        var doc = CreateDocumentWithParagraphs("Item 0", "Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", start },
            { "endParagraphIndex", end }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"paragraph {start} to {end}", result.Message);
    }

    #endregion

    #region Basic Convert Operations

    [Fact]
    public void Execute_ConvertsParagraphsToList()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted to list successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsRange()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("paragraph 0 to 2", result.Message);
    }

    [Fact]
    public void Execute_ReturnsConvertedCount()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Converted:", result.Message);
    }

    [Fact]
    public void Execute_ConvertsSingleParagraph()
    {
        var doc = CreateDocumentWithParagraphs("Single Item");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted to list successfully", result.Message);
    }

    #endregion

    #region List Type Parameter

    [Fact]
    public void Execute_DefaultListType_UsesBullet()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("List type: bullet", result.Message);
    }

    [Fact]
    public void Execute_WithNumberType_CreatesNumberList()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 },
            { "listType", "number" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("List type: number", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStartIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "endParagraphIndex", 1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutEndIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidStartIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 99 },
            { "endParagraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidEndIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeStartIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", -1 },
            { "endParagraphIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithStartGreaterThanEnd_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 2 },
            { "endParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("less than or equal", ex.Message);
    }

    #endregion
}
