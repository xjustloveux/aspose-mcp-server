using Aspose.Words;
using AsposeMcpServer.Handlers.Word.SectionBreak;
using AsposeMcpServer.Results.Word.SectionBreak;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.SectionBreak;

public class GetWordSectionsHandlerTests : WordHandlerTestBase
{
    private readonly GetWordSectionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
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

    #region Basic Get Sections Operations

    [Fact]
    public void Execute_ReturnsAllSections()
    {
        var doc = CreateDocumentWithMultipleSections(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsWordResult>(res);

        Assert.NotNull(result);
        Assert.Equal(3, result.TotalSections);
        Assert.NotNull(result.Sections);
    }

    [Fact]
    public void Execute_WithSectionIndex_ReturnsSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Section);
        Assert.Equal(1, result.Section.Index);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 10 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_ReturnsSectionBreakInfo()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Sections);
        Assert.NotEmpty(result.Sections);
        Assert.NotNull(result.Sections[0].SectionBreak);
        Assert.NotNull(result.Sections[0].SectionBreak.Type);
    }

    [Fact]
    public void Execute_ReturnsPageSetupInfo()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Sections);
        Assert.NotEmpty(result.Sections);
        Assert.NotNull(result.Sections[0].PageSetup);
        Assert.NotNull(result.Sections[0].PageSetup.PaperSize);
        Assert.NotNull(result.Sections[0].PageSetup.Orientation);
        Assert.NotNull(result.Sections[0].PageSetup.Margins);
    }

    [Fact]
    public void Execute_ReturnsContentStatistics()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Sections);
        Assert.NotEmpty(result.Sections);
        Assert.NotNull(result.Sections[0].ContentStatistics);
        Assert.True(result.Sections[0].ContentStatistics.Paragraphs >= 0);
        Assert.True(result.Sections[0].ContentStatistics.Tables >= 0);
    }

    #endregion
}
