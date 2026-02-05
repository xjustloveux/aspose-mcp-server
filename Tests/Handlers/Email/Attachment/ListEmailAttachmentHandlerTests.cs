using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Results.Email.Attachment;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Attachment;

/// <summary>
///     Tests for <see cref="ListEmailAttachmentHandler" />.
///     Verifies listing attachments from email files.
/// </summary>
public class ListEmailAttachmentHandlerTests : HandlerTestBase<object>
{
    private readonly ListEmailAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_List()
    {
        Assert.Equal("list", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file without attachments.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlFile(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Test",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates an EML file with the specified number of attachments.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="attachmentCount">The number of attachments to add.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlWithAttachments(string fileName, int attachmentCount)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Email With Attachments",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");

        for (var i = 0; i < attachmentCount; i++)
        {
            var attachFile = CreateTempFile(".txt", $"Attachment content {i}");
            message.Attachments.Add(new Aspose.Email.Attachment(attachFile));
        }

        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_WithNoAttachments_ReturnsEmptyList()
    {
        var path = CreateEmlFile("test_no_attachments.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetAttachmentsEmailResult>(result);
        var attachmentResult = (GetAttachmentsEmailResult)result;
        Assert.Equal(0, attachmentResult.Count);
        Assert.Empty(attachmentResult.Attachments);
        Assert.Contains("0 attachment(s)", attachmentResult.Message);
    }

    [Fact]
    public void Execute_WithOneAttachment_ReturnsSingleItem()
    {
        var path = CreateEmlWithAttachments("test_one_attachment.eml", 1);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetAttachmentsEmailResult>(result);
        var attachmentResult = (GetAttachmentsEmailResult)result;
        Assert.Equal(1, attachmentResult.Count);
        Assert.Single(attachmentResult.Attachments);
        Assert.Contains("1 attachment(s)", attachmentResult.Message);
    }

    [Fact]
    public void Execute_WithMultipleAttachments_ReturnsAll()
    {
        var path = CreateEmlWithAttachments("test_multiple_attachments.eml", 3);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetAttachmentsEmailResult>(result);
        var attachmentResult = (GetAttachmentsEmailResult)result;
        Assert.Equal(3, attachmentResult.Count);
        Assert.Equal(3, attachmentResult.Attachments.Count);
        Assert.Contains("3 attachment(s)", attachmentResult.Message);
    }

    [Fact]
    public void Execute_AttachmentInfo_HasCorrectProperties()
    {
        var path = CreateEmlWithAttachments("test_attachment_info.eml", 1);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var attachmentResult = (GetAttachmentsEmailResult)result;
        var attachment = attachmentResult.Attachments[0];
        Assert.Equal(0, attachment.Index);
        Assert.NotNull(attachment.Name);
        Assert.True(attachment.Size > 0);
    }

    [Fact]
    public void Execute_MultipleAttachments_HaveSequentialIndices()
    {
        var path = CreateEmlWithAttachments("test_sequential_indices.eml", 3);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", path }
        });

        var result = _handler.Execute(context, parameters);

        var attachmentResult = (GetAttachmentsEmailResult)result;
        for (var i = 0; i < attachmentResult.Attachments.Count; i++)
            Assert.Equal(i, attachmentResult.Attachments[i].Index);
    }

    #endregion
}
