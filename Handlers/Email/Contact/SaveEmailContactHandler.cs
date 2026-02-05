using Aspose.Email.Mapi;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Contact;

/// <summary>
///     Handler for saving (converting) an email contact to a different format.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SaveEmailContactHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "save";

    /// <summary>
    ///     Loads an email contact and saves it to a new file, optionally in a different format.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: path (source contact file), outputPath (destination file).
    ///     Optional: format ("vcf" or "msg", auto-detected from outputPath extension by default).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the contact was saved.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var format = parameters.GetOptional<string?>("format");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        var contact = LoadEmailContactHandler.LoadContact(path);

        var ext = format?.ToLowerInvariant() ?? Path.GetExtension(outputPath).ToLowerInvariant().TrimStart('.');
        contact.Save(outputPath, ext == "msg" ? ContactSaveFormat.Msg : ContactSaveFormat.VCard);

        return new SuccessResult
        {
            Message = $"Contact saved: {contact.NameInfo?.DisplayName ?? "Unknown"} -> {outputPath}"
        };
    }
}
