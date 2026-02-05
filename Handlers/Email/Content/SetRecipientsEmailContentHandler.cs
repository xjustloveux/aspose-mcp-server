using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for setting recipients of an email.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetRecipientsEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_recipients";

    /// <summary>
    ///     Loads an email file, sets its recipients, and saves it.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path).
    ///     Optional: from (sender address), to (comma-separated To addresses),
    ///     cc (comma-separated CC addresses), bcc (comma-separated BCC addresses),
    ///     outputPath (save location, defaults to path).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the recipients were set successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetOptional("outputPath", path);
        var from = parameters.GetOptional<string?>("from");
        var to = parameters.GetOptional<string?>("to");
        var cc = parameters.GetOptional<string?>("cc");
        var bcc = parameters.GetOptional<string?>("bcc");
        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        if (from != null)
            message.From = from;

        if (to != null)
        {
            message.To.Clear();
            foreach (var address in SplitAddresses(to))
                message.To.Add(address);
        }

        if (cc != null)
        {
            message.CC.Clear();
            foreach (var address in SplitAddresses(cc))
                message.CC.Add(address);
        }

        if (bcc != null)
        {
            message.Bcc.Clear();
            foreach (var address in SplitAddresses(bcc))
                message.Bcc.Add(address);
        }

        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Email recipients updated and saved to {outputPath}"
        };
    }

    /// <summary>
    ///     Splits a comma-separated string of email addresses into individual addresses.
    /// </summary>
    /// <param name="addresses">Comma-separated email addresses.</param>
    /// <returns>An enumerable of trimmed, non-empty email addresses.</returns>
    private static IEnumerable<string> SplitAddresses(string addresses)
    {
        return addresses.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(a => !string.IsNullOrWhiteSpace(a));
    }
}
