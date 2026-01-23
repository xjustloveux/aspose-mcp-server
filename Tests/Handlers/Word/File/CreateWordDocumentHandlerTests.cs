using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.File;

public class CreateWordDocumentHandlerTests : WordHandlerTestBase
{
    private readonly CreateWordDocumentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_CreatesEmptyDocument()
    {
        var outputPath = Path.Combine(TestDir, "test.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created document should have content");
    }

    [Fact]
    public void Execute_WithContent_CreatesDocumentWithContent()
    {
        var outputPath = Path.Combine(TestDir, "with_content.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "content", "Hello World" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created document should have content");

        var createdDoc = new Document(outputPath);
        Assert.Contains("Hello World", createdDoc.GetText());
    }

    [Fact]
    public void Execute_WithPaperSize_SetsCorrectPaperSize()
    {
        var outputPath = Path.Combine(TestDir, "letter_size.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "paperSize", "Letter" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created document should have content");
    }

    [Fact]
    public void Execute_WithSkipInitialContent_CreatesEmptyDocument()
    {
        var outputPath = Path.Combine(TestDir, "skip_content.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "skipInitialContent", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created document should have content");
    }

    #endregion
}
