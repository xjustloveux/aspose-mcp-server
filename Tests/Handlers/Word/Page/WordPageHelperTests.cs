using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class WordPageHelperTests : WordTestBase
{
    #region Helper Methods

    private static Document CreateDocumentWithSections(int sectionCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 1; i < sectionCount; i++)
        {
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write($"Section {i + 1}");
        }

        return doc;
    }

    #endregion

    #region GetTargetSections Tests - With Section Index

    [Fact]
    public void GetTargetSections_WithValidSectionIndex_ReturnsSingleSection()
    {
        var doc = CreateDocumentWithSections(3);

        var result = WordPageHelper.GetTargetSections(doc, 1, null);

        Assert.Single(result);
        Assert.Equal(1, result[0]);
    }

    [Fact]
    public void GetTargetSections_WithZeroIndex_ReturnsFirstSection()
    {
        var doc = CreateDocumentWithSections(3);

        var result = WordPageHelper.GetTargetSections(doc, 0, null);

        Assert.Single(result);
        Assert.Equal(0, result[0]);
    }

    [Fact]
    public void GetTargetSections_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithSections(2);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordPageHelper.GetTargetSections(doc, 10, null));

        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void GetTargetSections_WithNegativeSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithSections(2);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordPageHelper.GetTargetSections(doc, -1, null));

        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void GetTargetSections_WithInvalidIndexAndValidateFalse_ReturnsIndex()
    {
        var doc = CreateDocumentWithSections(2);

        var result = WordPageHelper.GetTargetSections(doc, 10, null, false);

        Assert.Single(result);
        Assert.Equal(10, result[0]);
    }

    #endregion

    #region GetTargetSections Tests - With Section Indices Array

    [Fact]
    public void GetTargetSections_WithSectionIndicesArray_ReturnsMultipleSections()
    {
        var doc = CreateDocumentWithSections(5);
        var indices = new JsonArray { 0, 2, 4 };

        var result = WordPageHelper.GetTargetSections(doc, null, indices);

        Assert.Equal(3, result.Count);
        Assert.Contains(0, result);
        Assert.Contains(2, result);
        Assert.Contains(4, result);
    }

    [Fact]
    public void GetTargetSections_WithEmptySectionIndicesArray_ReturnsAllSections()
    {
        var doc = CreateDocumentWithSections(3);
        var indices = new JsonArray();

        var result = WordPageHelper.GetTargetSections(doc, null, indices);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void GetTargetSections_WithInvalidIndexInArray_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithSections(2);
        var indices = new JsonArray { 0, 10 };

        var ex = Assert.Throws<ArgumentException>(() =>
            WordPageHelper.GetTargetSections(doc, null, indices));

        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void GetTargetSections_WithSectionIndicesArrayPrecedence_UsesSectionIndices()
    {
        var doc = CreateDocumentWithSections(5);
        var indices = new JsonArray { 1, 3 };

        var result = WordPageHelper.GetTargetSections(doc, 0, indices);

        Assert.Equal(2, result.Count);
        Assert.Contains(1, result);
        Assert.Contains(3, result);
    }

    #endregion

    #region GetTargetSections Tests - No Parameters

    [Fact]
    public void GetTargetSections_WithNoParameters_ReturnsAllSections()
    {
        var doc = CreateDocumentWithSections(4);

        var result = WordPageHelper.GetTargetSections(doc, null, null);

        Assert.Equal(4, result.Count);
        Assert.Equal([0, 1, 2, 3], result);
    }

    [Fact]
    public void GetTargetSections_WithSingleSection_ReturnsSingleSection()
    {
        var doc = new Document();

        var result = WordPageHelper.GetTargetSections(doc, null, null);

        Assert.Single(result);
        Assert.Equal(0, result[0]);
    }

    #endregion
}
