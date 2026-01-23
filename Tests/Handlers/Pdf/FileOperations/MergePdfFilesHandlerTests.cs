using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class MergePdfFilesHandlerTests : PdfHandlerTestBase
{
    private readonly MergePdfFilesHandler _handler = new();
    private readonly string _input1Path;
    private readonly string _input2Path;

    public MergePdfFilesHandlerTests()
    {
        _input1Path = Path.Combine(TestDir, "input1.pdf");
        using (var doc1 = new Document())
        {
            doc1.Pages.Add();
            doc1.Save(_input1Path);
        }

        _input2Path = Path.Combine(TestDir, "input2.pdf");
        using (var doc2 = new Document())
        {
            doc2.Pages.Add();
            doc2.Pages.Add();
            doc2.Save(_input2Path);
        }
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
    public void Execute_MergesPdfDocuments()
    {
        var outputPath = Path.Combine(TestDir, "merged.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "inputPaths", new[] { _input1Path, _input2Path } }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Merged file should have content");

        using var mergedDoc = new Document(outputPath);
        Assert.Equal(3, mergedDoc.Pages.Count);
    }

    #endregion

    #region Error Handling

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

    [Fact]
    public void Execute_WithoutInputPaths_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "merged.pdf") }
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
            { "outputPath", Path.Combine(TestDir, "merged.pdf") },
            { "inputPaths", Array.Empty<string>() }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
