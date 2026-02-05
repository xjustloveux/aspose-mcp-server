using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Attachment;

/// <summary>
///     Tests for <see cref="AddEmailAttachmentHandler" />.
///     Verifies adding attachments to email files.
/// </summary>
public class AddEmailAttachmentHandlerTests : HandlerTestBase<object>
{
    private readonly AddEmailAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
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
    ///     Creates an EML file with the specified number of existing attachments.
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
            var attachFile = CreateTempFile(".txt", $"Existing attachment {i}");
            message.Attachments.Add(new Aspose.Email.Attachment(attachFile));
        }

        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_AddsAttachmentToEmptyEmail()
    {
        var emailPath = CreateEmlFile("test_add.eml");
        var outputPath = CreateTestFilePath("test_add_output.eml");
        var attachmentFile = CreateTempFile(".txt", "Attachment content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("added successfully", success.Message);
        Assert.Contains("1 attachment(s)", success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Single(loaded.Attachments);
    }

    [Fact]
    public void Execute_AddsAttachmentToEmailWithExistingAttachments()
    {
        var emailPath = CreateEmlWithAttachments("test_add_existing.eml", 2);
        var outputPath = CreateTestFilePath("test_add_existing_output.eml");
        var attachmentFile = CreateTempFile(".txt", "New attachment content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("3 attachment(s)", success.Message);

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal(3, loaded.Attachments.Count);
    }

    [Fact]
    public void Execute_MessageContainsFileName()
    {
        var emailPath = CreateEmlFile("test_msg_filename.eml");
        var outputPath = CreateTestFilePath("test_msg_filename_output.eml");
        var attachmentFile = CreateTempFile(".txt", "Content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        var result = _handler.Execute(context, parameters);

        var success = (SuccessResult)result;
        Assert.Contains(Path.GetFileName(attachmentFile), success.Message);
    }

    [SkippableFact]
    public void Execute_OutputFileIsLoadable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var emailPath = CreateEmlFile("test_loadable.eml");
        var outputPath = CreateTestFilePath("test_loadable_output.eml");
        var attachmentFile = CreateTempFile(".txt", "Loadable content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        _handler.Execute(context, parameters);

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal("Test", loaded.Subject);
        Assert.Single(loaded.Attachments);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentEmailFile_ThrowsFileNotFoundException()
    {
        var outputPath = CreateTestFilePath("output_missing_email.eml");
        var attachmentFile = CreateTempFile(".txt", "Content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") },
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentAttachment_ThrowsFileNotFoundException()
    {
        var emailPath = CreateEmlFile("test_missing_attachment.eml");
        var outputPath = CreateTestFilePath("output_missing_attachment.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "attachmentPath", Path.Combine(TestDir, "nonexistent_attachment.txt") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var outputPath = CreateTestFilePath("output_no_path.eml");
        var attachmentFile = CreateTempFile(".txt", "Content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "attachmentPath", attachmentFile }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var emailPath = CreateEmlFile("test_no_output.eml");
        var attachmentFile = CreateTempFile(".txt", "Content");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "attachmentPath", attachmentFile }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutAttachmentPath_ThrowsArgumentException()
    {
        var emailPath = CreateEmlFile("test_no_attachment_path.eml");
        var outputPath = CreateTestFilePath("output_no_attachment_path.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
