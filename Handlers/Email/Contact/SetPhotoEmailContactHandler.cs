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
            throw new FileNotFoundException("The specified file was not found.");

        if (!File.Exists(photoPath))
            throw new FileNotFoundException("The specified file was not found.");

        var contact = GetEmailContactHandler.LoadContact(path);
        // B-2/H39: resolve symlinks immediately before the read and write sinks (bug 20260415-symlink-toctou-sweep).
        var resolvedPhotoPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(photoPath, context.ServerConfig?.AllowedBasePaths ?? [],
                "photoPath");
        var photoBytes = File.ReadAllBytes(resolvedPhotoPath);
        contact.Photo = new MapiContactPhoto(photoBytes, DetectPhotoImageFormat(photoBytes));

        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        var resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputPath");
        contact.Save(resolvedOutputPath, ext == ".msg" ? ContactSaveFormat.Msg : ContactSaveFormat.VCard);

        return new SuccessResult
        {
            Message = $"Photo set on contact: {contact.NameInfo?.DisplayName ?? "Unknown"} -> {outputPath}"
        };
    }

    /// <summary>
    ///     Detects the photo image format from the image content signature.
    ///     <see cref="MapiContactPhotoImageFormat" /> has no PNG member, so PNG (and any other
    ///     unrecognized signature) maps to <see cref="MapiContactPhotoImageFormat.Undefined" />
    ///     instead of being falsely declared JPEG.
    /// </summary>
    /// <param name="photoBytes">The photo file content.</param>
    /// <returns>The detected image format, or Undefined when the signature is not recognized.</returns>
    internal static MapiContactPhotoImageFormat DetectPhotoImageFormat(byte[] photoBytes)
    {
        return photoBytes switch
        {
            [0xFF, 0xD8, 0xFF, ..] => MapiContactPhotoImageFormat.Jpeg,
            [0x47, 0x49, 0x46, ..] => MapiContactPhotoImageFormat.Gif,
            [0x42, 0x4D, ..] => MapiContactPhotoImageFormat.Bmp,
            [0x49, 0x49, 0x2A, 0x00, ..] or [0x4D, 0x4D, 0x00, 0x2A, ..] => MapiContactPhotoImageFormat.Tiff,
            [0xD7, 0xCD, 0xC6, 0x9A, ..] => MapiContactPhotoImageFormat.Wmf,
            _ => MapiContactPhotoImageFormat.Undefined
        };
    }
}
