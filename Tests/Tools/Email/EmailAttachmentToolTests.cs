using Aspose.Email;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Email.Attachment;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Tools.Email;

/// <summary>
///     Integration tests for <see cref="EmailAttachmentTool" />.
///     Focuses on operation routing, file I/O, and parameter forwarding.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class EmailAttachmentToolTests : EmailTestBase
{
    private readonly EmailAttachmentTool _tool = new();

    #region Helper Methods

    /// <summary>
    ///     Creates an EML file with named attachments for testing.
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
            Subject = "Test Email",
            Body = "Test Body"
        };
        message.To.Add("recipient@example.com");

        foreach (var name in attachmentNames)
        {
            var attachDir = Path.Combine(TestDir, "attach_sources");
            Directory.CreateDirectory(attachDir);
            var attachFile = Path.Combine(attachDir, $"src_{Guid.NewGuid()}{Path.GetExtension(name)}");
            File.WriteAllText(attachFile, $"Content of {name}");
            TestFiles.Add(attachFile);
            var attachment = new Attachment(attachFile);
            attachment.Name = name;
            message.Attachments.Add(attachment);
        }

        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates a simple text file to be used as an attachment.
    /// </summary>
    /// <param name="fileName">The file name.</param>
    /// <param name="content">The file content.</param>
    /// <returns>The full path to the created file.</returns>
    private string CreateAttachmentFile(string fileName, string content = "Attachment content")
    {
        var filePath = CreateTestFilePath(fileName);
        File.WriteAllText(filePath, content);
        return filePath;
    }

    #endregion

    #region List Operation

    [Fact]
    public void List_WithNoAttachments_ShouldReturnEmptyList()
    {
        var emlPath = CreateEmlFile("test_list_empty.eml");

        var result = _tool.Execute("list", emlPath);

        var data = GetResultData<GetAttachmentsEmailResult>(result);
        Assert.Equal(0, data.Count);
        Assert.Empty(data.Attachments);
    }

    [Fact]
    public void List_WithAttachments_ShouldReturnAttachmentInfo()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_list.eml",
            "document.txt", "image.png");

        var result = _tool.Execute("list", emlPath);

        var data = GetResultData<GetAttachmentsEmailResult>(result);
        Assert.Equal(2, data.Count);
        Assert.Equal(2, data.Attachments.Count);
        Assert.Contains("2 attachment(s)", data.Message);
    }

    [Fact]
    public void List_AttachmentInfo_HasCorrectIndices()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_list_indices.eml",
            "first.txt", "second.txt", "third.txt");

        var result = _tool.Execute("list", emlPath);

        var data = GetResultData<GetAttachmentsEmailResult>(result);
        for (var i = 0; i < data.Attachments.Count; i++)
            Assert.Equal(i, data.Attachments[i].Index);
    }

    [Fact]
    public void List_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_list.eml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("list", fakePath));
    }

    #endregion

    #region Add Operation

    [Fact]
    public void Add_ShouldAddAttachmentToEmail()
    {
        var emlPath = CreateEmlFile("test_add_source.eml");
        var outputPath = CreateTestFilePath("test_add_output.eml");
        var attachmentFile = CreateAttachmentFile("new_attachment.txt");

        var result = _tool.Execute("add", emlPath,
            outputPath, attachmentFile);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added successfully", data.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Single(loaded.Attachments);
    }

    [Fact]
    public void Add_ToEmailWithExistingAttachments_ShouldAppend()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_add_existing.eml", "existing.txt");
        var outputPath = CreateTestFilePath("test_add_existing_output.eml");
        var attachmentFile = CreateAttachmentFile("new_file.txt");

        var result = _tool.Execute("add", emlPath,
            outputPath, attachmentFile);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("2 attachment(s)", data.Message);

        var loaded = MailMessage.Load(outputPath);
        Assert.Equal(2, loaded.Attachments.Count);
    }

    [Fact]
    public void Add_WithNonExistentEmail_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_add.eml");
        var outputPath = CreateTestFilePath("add_output.eml");
        var attachmentFile = CreateAttachmentFile("attach.txt");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", fakePath,
                outputPath, attachmentFile));
    }

    [Fact]
    public void Add_WithNonExistentAttachment_ShouldThrowFileNotFoundException()
    {
        var emlPath = CreateEmlFile("test_add_no_attach.eml");
        var outputPath = CreateTestFilePath("add_no_attach_output.eml");
        var fakeAttach = CreateTestFilePath("nonexistent_attachment.txt");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", emlPath,
                outputPath, fakeAttach));
    }

    #endregion

    #region Remove Operation

    [Fact]
    public void Remove_ShouldRemoveAttachmentByIndex()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_remove.eml",
            "first.txt", "second.txt");
        var outputPath = CreateTestFilePath("test_remove_output.eml");

        var result = _tool.Execute("remove", emlPath,
            outputPath, index: 0);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("removed successfully", data.Message);
        Assert.True(File.Exists(outputPath));

        var loaded = MailMessage.Load(outputPath);
        Assert.Single(loaded.Attachments);
    }

    [Fact]
    public void Remove_WithInvalidIndex_ShouldThrowArgumentOutOfRangeException()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_remove_oob.eml", "file.txt");
        var outputPath = CreateTestFilePath("test_remove_oob_output.eml");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _tool.Execute("remove", emlPath,
                outputPath, index: 5));
    }

    [Fact]
    public void Remove_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_remove.eml");
        var outputPath = CreateTestFilePath("remove_output.eml");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("remove", fakePath,
                outputPath, index: 0));
    }

    #endregion

    #region Extract Operation

    [Fact]
    public void Extract_ShouldExtractSingleAttachment()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_extract.eml",
            "document.txt", "image.txt");
        var outputDir = Path.Combine(TestDir, "extract_single");

        var result = _tool.Execute("extract", emlPath,
            outputDir: outputDir, index: 0);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("extracted to", data.Message);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void Extract_WithInvalidIndex_ShouldThrowArgumentOutOfRangeException()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_extract_oob.eml", "file.txt");
        var outputDir = Path.Combine(TestDir, "extract_oob");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _tool.Execute("extract", emlPath,
                outputDir: outputDir, index: 10));
    }

    [Fact]
    public void Extract_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_extract.eml");
        var outputDir = Path.Combine(TestDir, "extract_missing");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("extract", fakePath,
                outputDir: outputDir, index: 0));
    }

    #endregion

    #region ExtractAll Operation

    [Fact]
    public void ExtractAll_ShouldExtractAllAttachments()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_extract_all.eml",
            "alpha.txt", "beta.txt", "gamma.txt");
        var outputDir = Path.Combine(TestDir, "extract_all");

        var result = _tool.Execute("extract_all", emlPath,
            outputDir: outputDir);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("3 attachment(s)", data.Message);
        Assert.True(Directory.Exists(outputDir));

        var files = Directory.GetFiles(outputDir);
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public void ExtractAll_WithNoAttachments_ShouldReturnNoAttachmentsMessage()
    {
        var emlPath = CreateEmlFile("test_extract_all_empty.eml");
        var outputDir = Path.Combine(TestDir, "extract_all_empty");

        var result = _tool.Execute("extract_all", emlPath,
            outputDir: outputDir);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("No attachments found", data.Message);
    }

    [Fact]
    public void ExtractAll_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var fakePath = CreateTestFilePath("nonexistent_extract_all.eml");
        var outputDir = Path.Combine(TestDir, "extract_all_missing");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("extract_all", fakePath,
                outputDir: outputDir));
    }

    #endregion

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("test_unknown_op.eml");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", emlPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("LIST")]
    [InlineData("List")]
    [InlineData("list")]
    public void Execute_OperationShouldBeCaseInsensitive(string operation)
    {
        var emlPath = CreateEmlFile($"test_case_{operation}.eml");

        var result = _tool.Execute(operation, emlPath);

        var data = GetResultData<GetAttachmentsEmailResult>(result);
        Assert.Equal(0, data.Count);
    }

    #endregion

    #region Missing Required Parameters

    [Fact]
    public void Add_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("test_add_no_output.eml");
        var attachmentFile = CreateAttachmentFile("no_output_attach.txt");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", emlPath, attachmentPath: attachmentFile));
    }

    [Fact]
    public void Add_WithoutAttachmentPath_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlFile("test_add_no_attach.eml");
        var outputPath = CreateTestFilePath("no_attach_output.eml");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", emlPath, outputPath));
    }

    [Fact]
    public void Remove_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_remove_no_output.eml", "file.txt");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("remove", emlPath, index: 0));
    }

    [Fact]
    public void Extract_WithoutOutputDir_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_extract_no_dir.eml", "file.txt");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", emlPath, index: 0));
    }

    [Fact]
    public void ExtractAll_WithoutOutputDir_ShouldThrowArgumentException()
    {
        var emlPath = CreateEmlWithNamedAttachments("test_extract_all_no_dir.eml", "file.txt");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract_all", emlPath));
    }

    #endregion
}
