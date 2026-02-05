using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Content;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for retrieving email headers.
/// </summary>
[ResultType(typeof(EmailHeadersResult))]
public class GetHeadersEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_headers";

    /// <summary>
    ///     Loads an email file and returns all its headers.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to read).
    /// </param>
    /// <returns>An <see cref="EmailHeadersResult" /> containing all email headers.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);
        var headers = new List<EmailHeaderInfo>();

        for (var i = 0; i < message.Headers.Count; i++)
        {
            var key = message.Headers.GetKey(i);
            var value = message.Headers.Get(i);
            if (key != null)
                headers.Add(new EmailHeaderInfo
                {
                    Name = key,
                    Value = value ?? ""
                });
        }

        return new EmailHeadersResult
        {
            Headers = headers,
            Count = headers.Count,
            Message = $"Retrieved {headers.Count} header(s) from email"
        };
    }
}
