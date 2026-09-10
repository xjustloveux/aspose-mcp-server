using Aspose.Email;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;
using MailAttachment = Aspose.Email.Attachment;

namespace AsposeMcpServer.Tests.Handlers.Email;

/// <summary>
///     Covers the wiring half of RB-38. The two attachment-extract handlers write into a
///     caller-supplied output directory with bare <c>Directory.CreateDirectory</c> and
///     <c>attachment.Save</c> calls, so an output-side IO failure escaped as a raw BCL exception
///     whose message carries the absolute path. Every other family that writes one file per
///     extracted item (Excel, Word and PowerPoint OLE extract) routes the same failure through its
///     translator; Email was the only one that did not.
/// </summary>
public class EmailExtractErrorTranslationTests : TestBase
{
    /// <summary>Writes an email carrying one attachment.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The path written.</returns>
    private string CreateEmailWithAttachment(string fileName)
    {
        var attachmentPath = CreateTestFilePath("payload.txt");
        File.WriteAllText(attachmentPath, "payload");

        var path = CreateTestFilePath(fileName);
        using var message = new MailMessage();
        message.From = "sender@example.com";
        message.Subject = "with attachment";
        message.Body = "body";
        message.To.Add("recipient@example.com");
        message.Attachments.Add(new MailAttachment(attachmentPath));
        message.Save(path, SaveOptions.DefaultEml);
        return path;
    }

    /// <summary>
    ///     Builds an output directory path that cannot be created on any platform, because one of
    ///     its parent components is an existing regular file.
    /// </summary>
    /// <param name="blockerName">Name of the file that stands in the way.</param>
    /// <returns>The unusable directory path.</returns>
    private string CreateUnusableOutputDirectory(string blockerName)
    {
        var blocker = CreateTestFilePath(blockerName);
        File.WriteAllText(blocker, "not a directory");
        return Path.Combine(blocker, "attachments");
    }

    [Fact]
    public void ExtractAll_WhenTheOutputDirectoryCannotBeCreated_ShouldReportItAsAnOutputFailure()
    {
        var tool = new EmailAttachmentTool();
        var path = CreateEmailWithAttachment("extract_all_target.eml");

        var exception = Assert.ThrowsAny<Exception>(() =>
            tool.Execute("extract_all", path, outputDir: CreateUnusableOutputDirectory("blocker_all.txt")));

        Assert.Equal(ErrorMessageBuilder.OutputDirectoryNotWritable(), exception.Message);
        Assert.DoesNotContain(TestDir, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Extract_WhenTheOutputDirectoryCannotBeCreated_ShouldReportItAsAnOutputFailure()
    {
        var tool = new EmailAttachmentTool();
        var path = CreateEmailWithAttachment("extract_target.eml");

        var exception = Assert.ThrowsAny<Exception>(() =>
            tool.Execute("extract", path, outputDir: CreateUnusableOutputDirectory("blocker_one.txt"), index: 0));

        Assert.Equal(ErrorMessageBuilder.OutputDirectoryNotWritable(), exception.Message);
        Assert.DoesNotContain(TestDir, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ExtractAll_WithAWritableDirectory_ShouldStillExtract()
    {
        var tool = new EmailAttachmentTool();
        var path = CreateEmailWithAttachment("extract_ok.eml");
        var outputDir = Path.Combine(TestDir, "attachments_ok");

        tool.Execute("extract_all", path, outputDir: outputDir);

        Assert.Single(Directory.GetFiles(outputDir));
    }

    [Fact]
    public void Extract_WithAWritableDirectory_ShouldStillExtract()
    {
        var tool = new EmailAttachmentTool();
        var path = CreateEmailWithAttachment("extract_one_ok.eml");
        var outputDir = Path.Combine(TestDir, "attachment_one_ok");

        tool.Execute("extract", path, outputDir: outputDir, index: 0);

        Assert.Single(Directory.GetFiles(outputDir));
    }
}
