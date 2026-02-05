using Aspose.Email.Mapi;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Contact;

/// <summary>
///     Handler for setting a photo on an email contact.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetPhotoEmailContactHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_photo";

    /// <summary>
    ///     Loads an email contact, sets its photo from an image file, and saves to the output path.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: path (source contact file), outputPath (destination contact file), photoPath (image file).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the photo was set.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file or photo file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var photoPath = parameters.GetRequired<string>("photoPath");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
        SecurityHelper.ValidateFilePath(photoPath, "photoPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        if (!File.Exists(photoPath))
            throw new FileNotFoundException($"Photo file not found: {photoPath}");

        var contact = LoadEmailContactHandler.LoadContact(path);
        var photoBytes = File.ReadAllBytes(photoPath);
        contact.Photo = new MapiContactPhoto(photoBytes, MapiContactPhotoImageFormat.Jpeg);

        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        contact.Save(outputPath, ext == ".msg" ? ContactSaveFormat.Msg : ContactSaveFormat.VCard);

        return new SuccessResult
        {
            Message = $"Photo set on contact: {contact.NameInfo?.DisplayName ?? "Unknown"} -> {outputPath}"
        };
    }
}
