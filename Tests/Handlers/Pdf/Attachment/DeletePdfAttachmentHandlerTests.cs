using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Attachment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Attachment;

public class DeletePdfAttachmentHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Attachment Operations

    [SkippableFact]
    public void Execute_DeletesAttachment()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithAttachment("test.txt");
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentName", "test.txt" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNonExistentAttachment_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "attachmentName", "nonexistent.txt" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
