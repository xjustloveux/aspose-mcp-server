using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Attachment;

/// <summary>
///     Tests for <see cref="RemoveEmailAttachmentHandler" />.
///     Verifies removing attachments from email files by index.
/// </summary>
public class RemoveEmailAttachmentHandlerTests : HandlerTestBase<object>
{
    private readonly RemoveEmailAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Helper Methods

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

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_RemovesSingleAttachment()
    {
        var emailPath = CreateEmlWithAttachments("test_remove.eml", 1);
        var outputPath = CreateTestFilePath("test_remove_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("removed successfully", success.Message);
        Assert.Contains("0 attachment(s)", success.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Empty(loaded.Attachments);
    }

    [Fact]
    public void Execute_RemovesFirstOfMultipleAttachments()
    {
        var emailPath = CreateEmlWithAttachments("test_remove_first.eml", 3);
        var outputPath = CreateTestFilePath("test_remove_first_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("2 attachment(s)", success.Message);

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal(2, loaded.Attachments.Count);
    }

    [Fact]
    public void Execute_RemovesLastAttachment()
    {
        var emailPath = CreateEmlWithAttachments("test_remove_last.eml", 3);
        var outputPath = CreateTestFilePath("test_remove_last_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var loaded = MailMessage.Load(outputPath);
        Assert.Equal(2, loaded.Attachments.Count);
    }

    [Fact]
    public void Execute_MessageContainsRemovedAttachmentName()
    {
        var emailPath = CreateEmlWithAttachments("test_remove_name.eml", 1);
        var outputPath = CreateTestFilePath("test_remove_name_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var success = (SuccessResult)result;
        Assert.Contains("index 0", success.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlWithAttachments("test_negative_index.eml", 1);
        var outputPath = CreateTestFilePath("test_negative_index_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", -1 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlWithAttachments("test_oob_index.eml", 2);
        var outputPath = CreateTestFilePath("test_oob_index_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 5 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoAttachments_IndexZero_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlFile("test_no_attachments_remove.eml");
        var outputPath = CreateTestFilePath("test_no_attachments_remove_output.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputPath", outputPath },
            { "index", 0 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var outputPath = CreateTestFilePath("output_missing.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") },
            { "outputPath", outputPath },
            { "index", 0 }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Parameter Validation

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var outputPath = CreateTestFilePath("output_no_path.eml");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath },
            { "index", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var emailPath = CreateEmlWithAttachments("test_no_output.eml", 1);
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "index", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutIndex_ThrowsArgumentException()
    {
        var emailPath = CreateEmlWithAttachments("test_no_index.eml", 1);
        var outputPath = CreateTestFilePath("output_no_index.eml");
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
