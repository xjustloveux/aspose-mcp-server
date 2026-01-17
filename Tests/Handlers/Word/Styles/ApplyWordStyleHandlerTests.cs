using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Styles;

public class ApplyWordStyleHandlerTests : WordHandlerTestBase
{
    private readonly ApplyWordStyleHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ApplyStyle()
    {
        Assert.Equal("apply_style", _handler.Operation);
    }

    #endregion

    #region Various Styles

    [Theory]
    [InlineData("Heading 1")]
    [InlineData("Heading 2")]
    [InlineData("Normal")]
    [InlineData("Quote")]
    public void Execute_WithVariousStyles_AppliesCorrectly(string styleName)
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", styleName },
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("applied style", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTable(int rows, int cols)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < cols; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    #endregion

    #region Basic Apply Operations

    [Fact]
    public void Execute_AppliesStyleToParagraph()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" },
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("applied style", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithApplyToAllParagraphs_AppliesToAll()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Normal" },
            { "applyToAllParagraphs", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "NonExistentStyle" },
            { "paragraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoTargetSpecified_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("paragraphIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" },
            { "paragraphIndex", 0 },
            { "sectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sectionIndex", ex.Message);
    }

    #endregion

    #region Table Style Tests

    [Fact]
    public void Execute_WithTableIndex_AppliesStyleToTable()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Table Grid" },
            { "tableIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("applied style", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Table Grid" },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    #endregion

    #region Paragraph Indices Tests

    [Fact]
    public void Execute_WithParagraphIndices_AppliesToMultiple()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third", "Fourth");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Normal" },
            { "paragraphIndices", new[] { 0, 2 } } // NOSONAR CA1861 - Test data array
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndicesPartiallyValid_AppliesValidOnes()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Normal" },
            { "paragraphIndices", new[] { 0, 99 } } // NOSONAR CA1861 - Test data array
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1", result);
        AssertModified(context);
    }

    #endregion
}
