using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class SplitPdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly SplitPdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Split()
    {
        Assert.Equal("split", _handler.Operation);
    }

    #endregion

    #region Basic Split Operations

    [Fact]
    public void Execute_SplitsPdfDocument()
    {
        var outputDir = Path.Combine(TestDir, "split_output");
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("3", result);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(3, files.Length);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Split file {file} should have content");
        }
    }

    [Fact]
    public void Execute_WithPagesPerFile_SplitsPdfDocument()
    {
        var outputDir = Path.Combine(TestDir, "split_multi");
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "pagesPerFile", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("2", result);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(2, files.Length);
    }

    [SkippableFact]
    public void Execute_WithPageRange_ExtractsPages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode limits to 4 pages");

        var outputDir = Path.Combine(TestDir, "split_range");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "startPage", 2 },
            { "endPage", 4 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("extracted", result.ToLower());
        Assert.Contains("3", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputDir_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPagesPerFile_ThrowsArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "split_invalid");
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "pagesPerFile", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidStartPage_ThrowsArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "split_invalid_start");
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "startPage", 10 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
