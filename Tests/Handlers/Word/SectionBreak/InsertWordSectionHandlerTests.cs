using AsposeMcpServer.Handlers.Word.SectionBreak;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.SectionBreak;

public class InsertWordSectionHandlerTests : WordHandlerTestBase
{
    private readonly InsertWordSectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Insert()
    {
        Assert.Equal("insert", _handler.Operation);
    }

    #endregion

    #region Basic Insert Section Operations

    [Fact]
    public void Execute_InsertsNextPageSectionBreak()
    {
        var doc = CreateDocumentWithText("Some content.");
        var initialCount = doc.Sections.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionBreakType", "NextPage" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("section break inserted", result.ToLower());
        Assert.True(doc.Sections.Count > initialCount);
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsContinuousSectionBreak()
    {
        var doc = CreateDocumentWithText("Some content.");
        var initialCount = doc.Sections.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionBreakType", "Continuous" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("continuous", result.ToLower());
        Assert.True(doc.Sections.Count > initialCount);
    }

    [Fact]
    public void Execute_InsertsOddPageSectionBreak()
    {
        var doc = CreateDocumentWithText("Some content.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionBreakType", "OddPage" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("section break inserted", result.ToLower());
    }

    [Fact]
    public void Execute_InsertsEvenPageSectionBreak()
    {
        var doc = CreateDocumentWithText("Some content.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionBreakType", "EvenPage" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("section break inserted", result.ToLower());
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Some content.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionBreakType", "NextPage" },
            { "sectionIndex", 999 },
            { "insertAtParagraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
