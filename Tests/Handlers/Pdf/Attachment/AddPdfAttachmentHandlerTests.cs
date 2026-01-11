using AsposeMcpServer.Handlers.Pdf.Attachment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Attachment;

public class AddPdfAttachmentHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Attachment Operations

    [SkippableFact]
    public void Execute_AddsAttachment()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var tempFile = CreateTempFile(".txt", "Test attachment content");

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentPath", tempFile },
            { "attachmentName", "test.txt" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("test.txt", result);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AddsAttachmentWithDescription()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var tempFile = CreateTempFile(".txt", "Test content");

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentPath", tempFile },
            { "attachmentName", "document.txt" },
            { "description", "A test document" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentPath", "C:/nonexistent/file.txt" },
            { "attachmentName", "file.txt" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithDuplicateName_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var tempFile = CreateTempFile(".txt", "Test content");

        var context = CreateContext(document);
        var parameters1 = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentPath", tempFile },
            { "attachmentName", "duplicate.txt" }
        });

        _handler.Execute(context, parameters1);

        var parameters2 = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentPath", tempFile },
            { "attachmentName", "duplicate.txt" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters2));
    }

    #endregion
}
