using Aspose.Words;
using AsposeMcpServer.Handlers.Word.SectionBreak;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.SectionBreak;

public class DeleteWordSectionHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordSectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithMultipleSections(int sectionCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 0; i < sectionCount; i++)
        {
            builder.Write($"Section {i + 1} content");
            if (i < sectionCount - 1) builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        return doc;
    }

    #endregion

    #region Basic Delete Section Operations

    [Fact]
    public void Execute_DeletesSingleSection()
    {
        var doc = CreateDocumentWithMultipleSections(3);
        var initialCount = doc.Sections.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, doc.Sections.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesMultipleSections()
    {
        var doc = CreateDocumentWithMultipleSections(4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndices", new[] { 1, 2 } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted 2 section", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, doc.Sections.Count);
    }

    [Fact]
    public void Execute_WithSingleSection_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        Assert.Equal(1, doc.Sections.Count);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNeitherIndexProvided_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_CannotDeleteLastSection()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndices", new[] { 0, 1 } }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(1, doc.Sections.Count);
    }

    #endregion
}
