using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for retrieving the subject of an email.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class GetSubjectEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_subject";

    /// <summary>
    ///     Loads an email file and returns its subject line.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to read).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> containing the email subject in the message.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        return new SuccessResult
        {
            Message = message.Subject ?? ""
        };
    }
}
