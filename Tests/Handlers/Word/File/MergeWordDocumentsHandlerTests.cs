using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.File;

public class MergeWordDocumentsHandlerTests : WordHandlerTestBase
{
    private readonly MergeWordDocumentsHandler _handler = new();
    private readonly string _input1Path;
    private readonly string _input2Path;

    public MergeWordDocumentsHandlerTests()
    {
        _input1Path = Path.Combine(TestDir, "input1.docx");
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Write("Document 1 content");
        doc1.Save(_input1Path);

        _input2Path = Path.Combine(TestDir, "input2.docx");
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Write("Document 2 content");
        doc2.Save(_input2Path);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Merge()
    {
        Assert.Equal("merge", _handler.Operation);
    }

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesDocuments()
    {
        var outputPath = Path.Combine(TestDir, "merged.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { _input1Path, _input2Path } },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("merged", result.ToLower());
        Assert.Contains("2", result);
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        var mergedDoc = new Document(outputPath);
        Assert.True(mergedDoc.PageCount > 0, "Merged document should have pages");
        var text = mergedDoc.GetText();
        Assert.Contains("Document 1 content", text);
        Assert.Contains("Document 2 content", text);
    }

    [Fact]
    public void Execute_WithKeepSourceFormatting_MergesDocuments()
    {
        var outputPath = Path.Combine(TestDir, "merged_keep_source.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { _input1Path, _input2Path } },
            { "outputPath", outputPath },
            { "importFormatMode", "KeepSourceFormatting" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("merged", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        var mergedDoc = new Document(outputPath);
        Assert.True(mergedDoc.PageCount > 0, "Merged document should have pages");
        var text = mergedDoc.GetText();
        Assert.Contains("Document 1 content", text);
        Assert.Contains("Document 2 content", text);
    }

    [Fact]
    public void Execute_WithUnlinkHeadersFooters_MergesDocuments()
    {
        var outputPath = Path.Combine(TestDir, "merged_unlink.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { _input1Path, _input2Path } },
            { "outputPath", outputPath },
            { "unlinkHeadersFooters", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("merged", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        var mergedDoc = new Document(outputPath);
        Assert.True(mergedDoc.PageCount > 0, "Merged document should have pages");
        var text = mergedDoc.GetText();
        Assert.Contains("Document 1 content", text);
        Assert.Contains("Document 2 content", text);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutInputPaths_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "merged.docx") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyInputPaths_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", Array.Empty<string>() },
            { "outputPath", Path.Combine(TestDir, "merged.docx") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPaths", new[] { _input1Path, _input2Path } }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
