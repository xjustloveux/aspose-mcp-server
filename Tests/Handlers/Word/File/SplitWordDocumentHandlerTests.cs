using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.File;

public class SplitWordDocumentHandlerTests : WordHandlerTestBase
{
    private readonly SplitWordDocumentHandler _handler = new();
    private readonly string _inputPath;

    public SplitWordDocumentHandlerTests()
    {
        _inputPath = Path.Combine(TestDir, "input.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Section 1 content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Write("Section 2 content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Write("Section 3 content");
        doc.Save(_inputPath);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Split()
    {
        Assert.Equal("split", _handler.Operation);
    }

    #endregion

    #region Basic Split Operations

    [Fact]
    public void Execute_SplitsBySection()
    {
        var outputDir = Path.Combine(TestDir, "split_output");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("3", result);
        Assert.True(Directory.Exists(outputDir));

        var splitFiles = Directory.GetFiles(outputDir, "*.docx");
        Assert.Equal(3, splitFiles.Length);
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");

            var splitDoc = new Document(splitFile);
            Assert.True(splitDoc.PageCount > 0, "Split document should have at least one page");
        }
    }

    [Fact]
    public void Execute_SplitsBySectionExplicit()
    {
        var outputDir = Path.Combine(TestDir, "split_section");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDir", outputDir },
            { "splitBy", "section" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("section", result.ToLower());

        var splitFiles = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(splitFiles.Length > 0, "Split files should be created");
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");
        }
    }

    [Fact]
    public void Execute_SplitsByPage()
    {
        var outputDir = Path.Combine(TestDir, "split_page");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDir", outputDir },
            { "splitBy", "page" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("split", result.ToLower());
        Assert.Contains("page", result.ToLower());

        var splitFiles = Directory.GetFiles(outputDir, "*.docx");
        Assert.True(splitFiles.Length > 0, "Split files should be created");
        foreach (var splitFile in splitFiles)
        {
            var fileInfo = new FileInfo(splitFile);
            Assert.True(fileInfo.Length > 0, $"Split file {splitFile} should have content");

            var splitDoc = new Document(splitFile);
            Assert.True(splitDoc.PageCount > 0, "Split document should have at least one page");
        }
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrSessionId_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", Path.Combine(TestDir, "output") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputDir_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
