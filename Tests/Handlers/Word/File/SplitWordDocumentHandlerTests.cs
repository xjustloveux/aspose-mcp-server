using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("section", result.Message, StringComparison.OrdinalIgnoreCase);

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("page", result.Message, StringComparison.OrdinalIgnoreCase);

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

    #region Request Atomicity

    /// <summary>
    ///     R5-R01: a split published each section as it was produced, so one that failed on a later
    ///     section had already replaced the destinations of every section before it. The publish for
    ///     the second section is made to fail by putting a directory where its file has to go.
    /// </summary>
    /// <param name="splitBy">Which split mode to exercise; both fan out.</param>
    [Theory]
    [InlineData("section")]
    [InlineData("page")]
    public void Execute_WhenALaterOutputCannotBePublished_ShouldWriteNoneOfThem(string splitBy)
    {
        var outputDir = Path.Combine(TestDir, "atomic_" + splitBy);
        Directory.CreateDirectory(outputDir);

        var first = Path.Combine(outputDir, $"input_{splitBy}_1.docx");
        var blocked = Path.Combine(outputDir, $"input_{splitBy}_2.docx");
        Directory.CreateDirectory(blocked);

        var context = CreateContext(CreateEmptyDocument());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDir", outputDir },
            { "splitBy", splitBy }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));

        Assert.False(System.IO.File.Exists(first),
            "the first output was published before the second was known to fail");
        Assert.DoesNotContain(Directory.GetFiles(outputDir),
            f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal));
    }

    /// <summary>
    ///     §17.4.1: what a handled failure leaves behind when the output directory did not exist.
    ///     The tools promise that no output file is published, not that the filesystem is untouched
    ///     — the directory a request had to create can remain, and the description says so.
    /// </summary>
    [Fact]
    public void Execute_WhenTheOutputDirectoryIsNewAndTheRequestFails_ShouldPublishNoFiles()
    {
        var outputDir = Path.Combine(TestDir, "created_then_failed");
        Assert.False(Directory.Exists(outputDir));

        var context = CreateContext(CreateEmptyDocument());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDir", outputDir }
        });

        // The second section cannot be published, so none of them may be.
        Directory.CreateDirectory(Path.Combine(outputDir, "input_section_2.docx"));

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));

        Assert.Empty(Directory.GetFiles(outputDir));
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
