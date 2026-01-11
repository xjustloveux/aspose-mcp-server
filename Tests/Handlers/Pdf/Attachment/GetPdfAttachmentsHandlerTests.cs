using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Attachment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Attachment;

public class GetPdfAttachmentsHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfAttachmentsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Attachments Operations

    [Fact]
    public void Execute_ReturnsEmptyWhenNoAttachments()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No attachments found", result);
    }

    [SkippableFact]
    public void Execute_ReturnsAttachmentsList()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithAttachment("document.txt");
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 1", result);
        Assert.Contains("items", result);
    }

    [SkippableFact]
    public void Execute_ReturnsMultipleAttachments()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithMultipleAttachments();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 2", result);
    }

    #endregion

    #region Helper Methods

    private Document CreatePdfWithAttachment(string attachmentName)
    {
        var document = new Document();
        document.Pages.Add();

        var tempFile = CreateTempFile(".txt", "Test content");
        var fileSpec = new FileSpecification(tempFile, "Test attachment")
        {
            Name = attachmentName
        };
        document.EmbeddedFiles.Add(fileSpec);

        return document;
    }

    private Document CreatePdfWithMultipleAttachments()
    {
        var document = new Document();
        document.Pages.Add();

        var tempFile1 = CreateTempFile(".txt", "Content 1");
        var tempFile2 = CreateTempFile(".txt", "Content 2");

        var fileSpec1 = new FileSpecification(tempFile1, "First attachment")
        {
            Name = "file1.txt"
        };
        var fileSpec2 = new FileSpecification(tempFile2, "Second attachment")
        {
            Name = "file2.txt"
        };

        document.EmbeddedFiles.Add(fileSpec1);
        document.EmbeddedFiles.Add(fileSpec2);

        return document;
    }

    #endregion
}
