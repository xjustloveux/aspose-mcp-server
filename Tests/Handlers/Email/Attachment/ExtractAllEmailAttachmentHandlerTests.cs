using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Attachment;

/// <summary>
///     Tests for <see cref="ExtractAllEmailAttachmentHandler" />.
///     Verifies extracting all attachments from an email file.
/// </summary>
public class ExtractAllEmailAttachmentHandlerTests : HandlerTestBase<object>
{
    private readonly ExtractAllEmailAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ExtractAll()
    {
        Assert.Equal("extract_all", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var outputDir = Path.Combine(TestDir, "missing_file_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") },
            { "outputDir", outputDir }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
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
    ///     Creates an EML file with named attachments.
    /// </summary>
    /// <param name="fileName">The output file name.</param>
    /// <param name="attachmentNames">The names of attachments to add.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateEmlWithNamedAttachments(string fileName, params string[] attachmentNames)
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = "Email With Attachments",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");

        foreach (var name in attachmentNames)
        {
            var attachFile = CreateTempFile(Path.GetExtension(name), $"Content of {name}");
            var attachment = new Aspose.Email.Attachment(attachFile);
            attachment.Name = name;
            message.Attachments.Add(attachment);
        }

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
    public void Execute_WithNoAttachments_ReturnsNoAttachmentsMessage()
    {
        var emailPath = CreateEmlFile("test_no_attach.eml");
        var outputDir = Path.Combine(TestDir, "extract_all_empty");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("No attachments found", success.Message);
    }

    [Fact]
    public void Execute_WithSingleAttachment_ExtractsOne()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_extract_single.eml", "document.txt");
        var outputDir = Path.Combine(TestDir, "extract_all_single");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("1 attachment(s)", success.Message);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Execute_WithMultipleAttachments_ExtractsAll()
    {
        var emailPath = CreateEmlWithAttachments("test_extract_multiple.eml", 3);
        var outputDir = Path.Combine(TestDir, "extract_all_multiple");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("3 attachment(s)", success.Message);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public void Execute_CreatesOutputDirectoryIfNotExists()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_create_dir.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "new_extract_dir", "nested");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        _handler.Execute(context, parameters);

        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Execute_MessageContainsOutputDirectory()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_msg_dir.eml", "report.txt");
        var outputDir = Path.Combine(TestDir, "extract_msg_dir");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        var success = (SuccessResult)result;
        Assert.Contains(outputDir, success.Message);
    }

    [Fact]
    public void Execute_MessageContainsFileNames()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_filenames.eml",
            "alpha.txt", "beta.txt");
        var outputDir = Path.Combine(TestDir, "extract_filenames");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        var result = _handler.Execute(context, parameters);

        var success = (SuccessResult)result;
        Assert.Contains("alpha.txt", success.Message);
        Assert.Contains("beta.txt", success.Message);
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "no_path_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputDir_ThrowsArgumentException()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_no_outdir.eml", "file.txt");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
