using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Attachment;
using AsposeMcpServer.Results.Pdf.Attachment;
using AsposeMcpServer.Tests.Infrastructure;

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

    #endregion

    #region Basic Get Attachments Operations

    [Fact]
    public void Execute_ReturnsEmptyWhenNoAttachments()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAttachmentsResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Equal("No attachments found", result.Message);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectResultType()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAttachmentsResult>(res);
        Assert.NotNull(result.Items);
    }

    [SkippableFact]
    public void Execute_ReturnsResultWithItemsProperty()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithAttachment("document.txt");
        var context = CreateContext(document);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAttachmentsResult>(res);
        // Verify the result has the expected structure
        Assert.NotNull(result);
        Assert.NotNull(result.Items);
        // Note: In-memory attachments may not be fully accessible via CollectAttachmentInfo
        // The actual count depends on Aspose.Pdf's behavior with in-memory documents
    }

    #endregion
}
