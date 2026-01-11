using Aspose.Words;
using AsposeMcpServer.Handlers.Word.MailMerge;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Mail merge completed", result);
        Assert.Contains("Fields merged: 2", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains(outputPath, result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Mail merge completed", result);
        Assert.Contains("multiple records", result);
        Assert.Contains("Records processed: 2", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Output files:", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Mail merge completed", result);
        Assert.Contains("Cleanup applied", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Mail merge completed", result);
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
}
