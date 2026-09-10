using Aspose.Email;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Email;

namespace AsposeMcpServer.Tests.Handlers.Email;

/// <summary>
///     Covers LOW-07 (an address list split on every comma) and LOW-08 (no CR/LF policy on
///     header inputs).
///     <para>
///         Measured behaviour that motivates these tests: assigning
///         <c>"Hello&lt;CR&gt;&lt;LF&gt;Bcc: attacker@example.com"</c> to
///         <see cref="MailMessage.Subject" /> and saving as EML writes a genuine
///         <c>Bcc: attacker@example.com</c> header into the file. Aspose passes the line break
///         through, so the caller decides who else receives the message.
///     </para>
/// </summary>
public class EmailHeaderInjectionTests : TestBase
{
    /// <summary>Creates a minimal EML file to operate on.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The path written.</returns>
    private string CreateEmail(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        using var message = new MailMessage();
        message.From = "sender@example.com";
        message.Subject = "Original";
        message.Body = "body";
        message.To.Add("first@example.com");
        message.Save(path, SaveOptions.DefaultEml);
        return path;
    }

    [Fact]
    public void SetRecipients_WithCommaInsideQuotedDisplayName_ShouldKeepOneRecipient()
    {
        var tool = new EmailContentTool();
        var path = CreateEmail("quoted.eml");
        var outputPath = CreateTestFilePath("quoted_out.eml");

        tool.Execute("set_recipients", path, outputPath,
            to: "\"Last, First\" <user@example.com>, second@example.com");

        using var saved = MailMessage.Load(outputPath);
        Assert.Equal(2, saved.To.Count);
        Assert.Equal("user@example.com", saved.To[0].Address);
        Assert.Equal("Last, First", saved.To[0].DisplayName);
        Assert.Equal("second@example.com", saved.To[1].Address);
    }

    [Fact]
    public void SetSubject_WithLineBreak_ShouldBeRefused()
    {
        var tool = new EmailContentTool();
        var path = CreateEmail("subject.eml");
        var outputPath = CreateTestFilePath("subject_out.eml");

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("set_subject", path, outputPath,
                subject: "Hello\r\nBcc: attacker@example.com"));

        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaders_WithLineBreakInValue_ShouldBeRefused()
    {
        var tool = new EmailContentTool();
        var path = CreateEmail("header.eml");
        var outputPath = CreateTestFilePath("header_out.eml");

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("set_headers", path, outputPath,
                name: "X-Custom", value: "ok\r\nBcc: attacker@example.com"));

        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void SetRecipients_WithLineBreakInAddressList_ShouldBeRefused()
    {
        var tool = new EmailContentTool();
        var path = CreateEmail("recipients.eml");
        var outputPath = CreateTestFilePath("recipients_out.eml");

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("set_recipients", path, outputPath,
                to: "user@example.com\r\nBcc: attacker@example.com"));

        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void CreateEmail_WithLineBreakInSubject_ShouldBeRefused()
    {
        var tool = new EmailFileTool();
        var outputPath = CreateTestFilePath("created.eml");

        Assert.ThrowsAny<ArgumentException>(() =>
            tool.Execute("create", outputPath: outputPath,
                subject: "Hello\r\nBcc: attacker@example.com", body: "body"));

        Assert.False(File.Exists(outputPath));
    }

    [SkippableFact]
    public void SetSubject_WithOrdinarySubject_ShouldStillWork()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "Evaluation mode appends watermark to subject");
        var tool = new EmailContentTool();
        var path = CreateEmail("plain.eml");
        var outputPath = CreateTestFilePath("plain_out.eml");

        tool.Execute("set_subject", path, outputPath, subject: "Quarterly report");

        using var saved = MailMessage.Load(outputPath);
        Assert.Equal("Quarterly report", saved.Subject);
    }

    /// <summary>
    ///     The shared string and recipient limits apply to an address list too.
    ///     <para>
    ///         The splitter walks the value character by character and materialises one entry per
    ///         address, so an unbounded header value set both the parsing work and the number of
    ///         recipients that would be handed to the mail library (R2-S08).
    ///     </para>
    /// </summary>
    [Fact]
    public void Split_WithAnOverlongList_ShouldBeRejected()
    {
        var addresses = new string('a', 10_001);

        Assert.Throws<ArgumentException>(() => EmailAddressListHelper.Split(addresses));
    }

    [Fact]
    public void Split_WithTooManyRecipients_ShouldBeRejected()
    {
        var addresses = string.Join(",", Enumerable.Range(0, 1_001).Select(i => $"u{i}@e.co"));

        Assert.Throws<ArgumentException>(() => EmailAddressListHelper.Split(addresses));
    }

    [Fact]
    public void Split_WithAnOrdinaryList_ShouldBeAccepted()
    {
        var result = EmailAddressListHelper.Split("a@b.com, \"Last, First\" <c@d.com>");

        Assert.Equal(2, result.Count);
    }
}
