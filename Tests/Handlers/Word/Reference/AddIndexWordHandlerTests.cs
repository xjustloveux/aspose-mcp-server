using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Reference;

public class AddIndexWordHandlerTests : WordHandlerTestBase
{
    private readonly AddIndexWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddIndex()
    {
        Assert.Equal("add_index", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsIndexEntries()
    {
        var doc = CreateDocumentWithText("Sample document text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "indexEntries", "[{\"text\":\"Sample\"},{\"text\":\"Document\"}]" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var xeFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldIndexEntry).ToList();
        var indexFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldIndex).ToList();
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Equal(2, xeFields.Count);
        Assert.NotEmpty(indexFields);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSubEntries_AddsIndexWithSubEntries()
    {
        var doc = CreateDocumentWithText("Sample document.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "indexEntries", "[{\"text\":\"Animals\",\"subEntry\":\"Dogs\"}]" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var xeFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldIndexEntry).ToList();
        Assert.NotEmpty(xeFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains("Animals", xeFields[0].GetFieldCode());
            Assert.Contains("Dogs", xeFields[0].GetFieldCode());
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutIndexEntries_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidJson_ThrowsException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "indexEntries", "invalid json" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Reported Counts And Limits

    /// <summary>
    ///     The result reported the length of the array, but an element that is not an object, or
    ///     carries no usable text, inserts nothing (R3-C06).
    /// </summary>
    [Fact]
    public void Execute_WithEntriesThatInsertNothing_ShouldReportOnlyWhatItInserted()
    {
        var doc = CreateDocumentWithText("Sample document text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            {
                "indexEntries",
                "[{\"text\":\"Kept\"},\"a bare string\",42,{\"text\":\"\"},{\"subEntry\":\"no text\"},{\"text\":7}]"
            }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Inserted: 1", result.Message);
        Assert.Contains("skipped 5", result.Message);

        var xeFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldIndexEntry).ToList();
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Single(xeFields);
    }

    /// <summary>
    ///     The entries arrive as a JSON string, so the bound on array parameters never saw them;
    ///     the count only exists after decoding (R3-R05).
    /// </summary>
    [Fact]
    public void Execute_WithMoreEntriesThanAllowed_ShouldBeRefusedBeforeInserting()
    {
        var doc = CreateDocumentWithText("Sample document text.");
        var context = CreateContext(doc);
        var entries = string.Join(",", Enumerable.Range(0, 1001).Select(i => $"{{\"text\":\"e{i}\"}}"));
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "indexEntries", "[" + entries + "]" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.DoesNotContain(doc.Range.Fields, f => f.Type == FieldType.FieldIndexEntry);
    }

    [Fact]
    public void Execute_WithExactlyTheAllowedNumberOfEntries_ShouldStillInsert()
    {
        var doc = CreateDocumentWithText("Sample document text.");
        var context = CreateContext(doc);
        var entries = string.Join(",", Enumerable.Range(0, 1000).Select(i => $"{{\"text\":\"e{i}\"}}"));
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "indexEntries", "[" + entries + "]" },
            { "insertIndexAtEnd", false }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Inserted: 1000", result.Message);
    }

    #endregion
}
