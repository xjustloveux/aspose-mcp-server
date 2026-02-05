using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.FileOperations;

/// <summary>
///     Handler for creating a new email file.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CreateEmailFileHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new email file with the specified properties.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: outputPath (file path to save the new email).
    ///     Optional: subject, body, from, to, isHtml (default: false).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the email was created successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when outputPath is missing or invalid.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var message = new MailMessage
        {
            From = p.From ?? "noreply@example.com",
            Subject = p.Subject ?? ""
        };

        if (p.IsHtml)
            message.HtmlBody = p.Body ?? "";
        else
            message.Body = p.Body ?? "";

        if (!string.IsNullOrEmpty(p.To))
            message.To.Add(p.To);

        var saveOptions = DetectSaveOptions(p.OutputPath);
        message.Save(p.OutputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Email created and saved to {p.OutputPath}"
        };
    }

    /// <summary>
    ///     Detects the appropriate save options based on file extension.
    /// </summary>
    /// <param name="path">The output file path.</param>
    /// <returns>The save options for the detected format.</returns>
    internal static SaveOptions DetectSaveOptions(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".msg" => SaveOptions.DefaultMsgUnicode,
            ".mhtml" or ".mht" => SaveOptions.DefaultMhtml,
            ".html" or ".htm" => SaveOptions.DefaultHtml,
            _ => SaveOptions.DefaultEml
        };
    }

    /// <summary>
    ///     Extracts create parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static CreateParameters ExtractParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetRequired<string>("outputPath"),
            parameters.GetOptional<string?>("subject"),
            parameters.GetOptional<string?>("body"),
            parameters.GetOptional<string?>("from"),
            parameters.GetOptional<string?>("to"),
            parameters.GetOptional("isHtml", false));
    }

    /// <summary>
    ///     Parameters for the create email operation.
    /// </summary>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="Subject">The email subject.</param>
    /// <param name="Body">The email body text.</param>
    /// <param name="From">The sender email address.</param>
    /// <param name="To">The recipient email address.</param>
    /// <param name="IsHtml">Whether the body is HTML content.</param>
    private sealed record CreateParameters(
        string OutputPath,
        string? Subject,
        string? Body,
        string? From,
        string? To,
        bool IsHtml);
}
