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
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing or invalid, or when the change would address
    ///     more recipients than one message may carry.
    /// </exception>
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
        path = SecurityHelper.ResolveAndEnsureWithinAllowlist(path,
            context.ServerConfig?.AllowedBasePaths ?? [], "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        EmailAddressListHelper.EnsureNoHeaderInjection(from, "from");
        EmailAddressListHelper.EnsureNoHeaderInjection(to, "to");
        EmailAddressListHelper.EnsureNoHeaderInjection(cc, "cc");
        EmailAddressListHelper.EnsureNoHeaderInjection(bcc, "bcc");

        if (!File.Exists(path))
            throw new FileNotFoundException("The specified file was not found.");

        using var message = MailMessage.Load(path);

        if (from != null)
            message.From = from;

        // Each field was bounded on its own, so three full fields addressed three times the
        // limit (R3-C08). The total is counted over the message as it would be afterwards — a
        // field the caller did not supply keeps the addresses it already has — and it is counted
        // before anything is cleared, so a refused call leaves the message untouched.
        var newTo = to != null ? SplitAddresses(to) : null;
        var newCc = cc != null ? SplitAddresses(cc) : null;
        var newBcc = bcc != null ? SplitAddresses(bcc) : null;

        EmailAddressListHelper.EnsureRecipientTotal(
            (newTo?.Count ?? message.To.Count)
            + (newCc?.Count ?? message.CC.Count)
            + (newBcc?.Count ?? message.Bcc.Count));

        if (newTo != null)
        {
            message.To.Clear();
            foreach (var address in newTo)
                message.To.Add(address);
        }

        if (newCc != null)
        {
            message.CC.Clear();
            foreach (var address in newCc)
                message.CC.Add(address);
        }

        if (newBcc != null)
        {
            message.Bcc.Clear();
            foreach (var address in newBcc)
                message.Bcc.Add(address);
        }

        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        // H38: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputPath));
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Email recipients updated and saved to {outputPath}"
        };
    }

    /// <summary>
    ///     Splits an address list into individual addresses, honouring quoted display names.
    /// </summary>
    /// <param name="addresses">The address list, e.g. <c>"Last, First" &lt;a@b.com&gt;, c@d.com</c>.</param>
    /// <returns>The individual address entries, trimmed, with empty entries removed.</returns>
    private static IReadOnlyList<string> SplitAddresses(string addresses)
    {
        return EmailAddressListHelper.Split(addresses);
    }
}
