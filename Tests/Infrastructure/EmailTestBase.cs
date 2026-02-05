using Aspose.Email;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for Email tool tests providing Email-specific functionality.
///     Email tools do not use DocumentContext/Session — they operate directly on files.
/// </summary>
public abstract class EmailTestBase : TestBase
{
    /// <summary>
    ///     Creates a simple EML email file for testing.
    /// </summary>
    /// <param name="fileName">The file name to create.</param>
    /// <param name="subject">The email subject.</param>
    /// <param name="body">The email body text.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateEmlFile(string fileName, string subject = "Test Subject", string body = "Test Body")
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            Body = body
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultEml);
        return filePath;
    }

    /// <summary>
    ///     Creates a simple MSG email file for testing.
    /// </summary>
    /// <param name="fileName">The file name to create.</param>
    /// <param name="subject">The email subject.</param>
    /// <param name="body">The email body text.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateMsgFile(string fileName, string subject = "Test Subject", string body = "Test Body")
    {
        var filePath = CreateTestFilePath(fileName);
        var message = new MailMessage
        {
            From = "sender@example.com",
            Subject = subject,
            Body = body
        };
        message.To.Add("recipient@example.com");
        message.Save(filePath, SaveOptions.DefaultMsgUnicode);
        return filePath;
    }

    /// <summary>
    ///     Checks if Aspose.Email is running in evaluation mode.
    /// </summary>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Email)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}
