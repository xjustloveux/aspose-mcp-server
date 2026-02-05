using Aspose.Email;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Attachment;

/// <summary>
///     Tests for <see cref="ExtractEmailAttachmentHandler" />.
///     Verifies extracting a specific attachment by index from an email file.
/// </summary>
public class ExtractEmailAttachmentHandlerTests : HandlerTestBase<object>
{
    private readonly ExtractEmailAttachmentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file with the specified number of attachments.
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
    public void Execute_ExtractsAttachmentToDirectory()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_extract.eml", "document.txt");
        var outputDir = Path.Combine(TestDir, "extract_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        var success = (SuccessResult)result;
        Assert.Contains("extracted to", success.Message);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Execute_ExtractsSpecificAttachmentByIndex()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_extract_specific.eml",
            "first.txt", "second.txt", "third.txt");
        var outputDir = Path.Combine(TestDir, "extract_specific_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Execute_CreatesOutputDirectoryIfNotExists()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_create_dir.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "new_subdir", "nested");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Execute_MessageContainsAttachmentName()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_msg_name.eml", "report.txt");
        var outputDir = Path.Combine(TestDir, "extract_msg_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var success = (SuccessResult)result;
        Assert.Contains("report.txt", success.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_neg_idx.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "neg_idx_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", -1 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_oob_idx.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "oob_idx_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 5 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoAttachments_ThrowsArgumentOutOfRangeException()
    {
        var emailPath = CreateEmlFile("test_no_attach_extract.eml");
        var outputDir = Path.Combine(TestDir, "no_attach_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir },
            { "index", 0 }
        });

        Assert.Throws<ArgumentOutOfRangeException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var outputDir = Path.Combine(TestDir, "missing_file_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.eml") },
            { "outputDir", outputDir },
            { "index", 0 }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
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
            { "outputDir", outputDir },
            { "index", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputDir_ThrowsArgumentException()
    {
        var emailPath = CreateEmlWithNamedAttachments("test_no_outputdir.eml", "file.txt");
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
        var emailPath = CreateEmlWithNamedAttachments("test_no_index.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "no_index_output");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", emailPath },
            { "outputDir", outputDir }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
