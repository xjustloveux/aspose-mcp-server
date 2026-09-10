using Aspose.Words;
using AsposeMcpServer.Handlers.Word.MailMerge;
using AsposeMcpServer.Results.Word.MailMerge;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.MailMerge;

public class ExecuteMailMergeHandlerTests : WordHandlerTestBase
{
    private readonly ExecuteMailMergeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Execute()
    {
        Assert.Equal("execute", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateTemplateDocument()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD Name");
        builder.Write(" from ");
        builder.InsertField("MERGEFIELD Company");
        builder.Writeln("!");
        return doc;
    }

    #endregion

    #region Request Atomicity

    /// <summary>
    ///     R4-R02: a merge over several records used to publish each document the moment it was
    ///     produced, so one that failed on a later record had already replaced the destinations of
    ///     every record before it — and a caller who retried could not tell which files belonged to
    ///     the run that failed. The publish for the second record is made to fail by putting a
    ///     directory where its file has to go.
    /// </summary>
    [Fact]
    public void Execute_WhenALaterRecordCannotBePublished_ShouldWriteNoneOfThem()
    {
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var outputPath = Path.Combine(TestDir, "atomic.docx");
        var first = Path.Combine(TestDir, "atomic_1.docx");
        var blocked = Path.Combine(TestDir, "atomic_2.docx");
        Directory.CreateDirectory(blocked);

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[{\"Name\": \"One\"}, {\"Name\": \"Two\"}]" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));

        Assert.False(System.IO.File.Exists(first),
            "the first record was published before the second was known to fail");
        Assert.DoesNotContain(Directory.GetFiles(TestDir),
            f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal));
    }

    [Fact]
    public void Execute_WithSeveralRecords_ShouldPublishAllOfThemAndLeaveNoStaging()
    {
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var outputPath = Path.Combine(TestDir, "batch.docx");

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[{\"Name\": \"One\"}, {\"Name\": \"Two\"}, {\"Name\": \"Three\"}]" }
        });

        var result = Assert.IsType<MailMergeResult>(_handler.Execute(context, parameters));

        Assert.Equal(3, result.RecordsProcessed);
        foreach (var file in result.OutputFiles)
            Assert.True(System.IO.File.Exists(file), $"{file} was reported but not written");
        Assert.DoesNotContain(Directory.GetFiles(TestDir),
            f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal));
    }

    #endregion

    #region Result Properties

    [Fact]
    public void Execute_ReturnsCorrectProperties()
    {
        var outputPath = Path.Combine(TestDir, "properties_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"John\", \"Company\": \"Corp\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.NotEmpty(result.TemplateSource);
        Assert.Equal(2, result.FieldsMerged);
        Assert.Equal(1, result.RecordsProcessed);
        Assert.Single(result.OutputFiles);
        Assert.Equal(outputPath, result.OutputFiles[0]);
    }

    [Fact]
    public void Execute_WithExtraDataKeys_ReportsOnlyMatchedFields()
    {
        var outputPath = Path.Combine(TestDir, "extra_keys_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\", \"Company\": \"Corp\", \"ExtraField\": \"Ignored\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(2, result.FieldsMerged);
    }

    [Fact]
    public void Execute_WithPartialData_ReportsOnlyMatchedFields()
    {
        var outputPath = Path.Combine(TestDir, "partial_data_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"OnlyName\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(1, result.FieldsMerged);
    }

    [Fact]
    public void Execute_WithNoMatchingFields_ReportsZeroFieldsMerged()
    {
        var outputPath = Path.Combine(TestDir, "no_match_output.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello World");
        var context = CreateContext(doc);

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\", \"Company\": \"Corp\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(0, result.FieldsMerged);
    }

    [Fact]
    public void Execute_WithMultipleRecords_ReportsOnlyMatchedFields()
    {
        var outputPath = Path.Combine(TestDir, "multi_match_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[{\"Name\": \"John\", \"Extra\": \"X\"}, {\"Name\": \"Jane\", \"Company\": \"Corp\"}]" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(2, result.FieldsMerged);
    }

    #endregion

    #region Single Record Mail Merge

    [Fact]
    public void Execute_WithSingleRecord_MergesFields()
    {
        var outputPath = Path.Combine(TestDir, "merged_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"John Doe\", \"Company\": \"Acme Corp\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(2, result.FieldsMerged);
        Assert.Equal(1, result.RecordsProcessed);
        Assert.Single(result.OutputFiles);
        Assert.True(System.IO.File.Exists(outputPath));

        var mergedDoc = new Document(outputPath);
        var text = mergedDoc.GetText();
        Assert.Contains("John Doe", text);
        Assert.Contains("Acme Corp", text);
    }

    [Fact]
    public void Execute_WithSingleRecord_ReturnsOutputPath()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Contains(outputPath, result.OutputFiles);
    }

    #endregion

    #region Multiple Records Mail Merge

    [Fact]
    public void Execute_WithMultipleRecords_CreatesMultipleFiles()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[{\"Name\": \"John\"}, {\"Name\": \"Jane\"}]" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.Equal(2, result.RecordsProcessed);
        Assert.Equal(2, result.OutputFiles.Count);
    }

    [Fact]
    public void Execute_WithMultipleRecords_ReturnsOutputFiles()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[{\"Name\": \"John\"}, {\"Name\": \"Jane\"}]" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.NotEmpty(result.OutputFiles);
        Assert.True(result.OutputFiles.Count > 0);
    }

    #endregion

    #region Cleanup Options

    [Fact]
    public void Execute_WithCleanupOptions_AppliesCleanup()
    {
        var outputPath = Path.Combine(TestDir, "cleaned_output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\"}" },
            { "cleanupOptions", "RemoveUnusedFields,RemoveEmptyParagraphs" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        Assert.NotNull(result.CleanupApplied);
        Assert.Contains("RemoveUnusedFields", result.CleanupApplied);
    }

    [Fact]
    public void Execute_WithoutCleanupOptions_UsesDefaultCleanup()
    {
        var outputPath = Path.Combine(TestDir, "default_cleanup.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<MailMergeResult>(res);

        // Default cleanup options are applied
        Assert.NotNull(result.CleanupApplied);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", "{\"Name\": \"Test\"}" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("outputPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutDataOrDataArray_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("data", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithBothDataAndDataArray_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "data", "{\"Name\": \"Test\"}" },
            { "dataArray", "[{\"Name\": \"John\"}]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Cannot specify both", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyDataArray_ThrowsArgumentException()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[]" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("No data provided", ex.Message);
    }

    #endregion

    #region Record Limits

    /// <summary>
    ///     The records arrive as a JSON string, so the bound on array parameters never saw them.
    ///     Each record clones the whole document and writes a file (R3-R05).
    /// </summary>
    [Fact]
    public void Execute_WithMoreRecordsThanAllowed_ShouldBeRefusedBeforeCloning()
    {
        var outputPath = Path.Combine(TestDir, "too_many_records.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var records = string.Join(",",
            Enumerable.Range(0, 1001).Select(i => $"{{\"Name\": \"n{i}\"}}"));
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[" + records + "]" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Empty(Directory.GetFiles(TestDir, "too_many_records*"));
    }

    [Fact]
    public void Execute_WithTheAllowedNumberOfRecords_ShouldStillMerge()
    {
        var outputPath = Path.Combine(TestDir, "allowed_records.docx");
        var doc = CreateTemplateDocument();
        var context = CreateContext(doc);
        var records = string.Join(",",
            Enumerable.Range(0, 5).Select(i => $"{{\"Name\": \"n{i}\"}}"));
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "dataArray", "[" + records + "]" }
        });

        var result = Assert.IsType<MailMergeResult>(_handler.Execute(context, parameters));

        Assert.Equal(5, result.RecordsProcessed);
    }

    #endregion
}
