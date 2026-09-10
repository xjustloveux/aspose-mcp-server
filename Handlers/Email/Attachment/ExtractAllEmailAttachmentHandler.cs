using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Email;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for extracting all attachments from an email message.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractAllEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "extract_all";

    /// <summary>
    ///     Extracts all attachments from the email and saves them to the output directory.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), outputDir (output directory).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming all attachments were extracted.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputDir = parameters.GetRequired<string>("outputDir");

        SecurityHelper.ValidateFilePath(path, "path", true);
        path = SecurityHelper.ResolveAndEnsureWithinAllowlist(path,
            context.ServerConfig?.AllowedBasePaths ?? [], "path");
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        if (!File.Exists(path))
            throw new FileNotFoundException("The specified file was not found.");

        using var message = MailMessage.Load(path);

        if (message.Attachments.Count == 0)
            return new SuccessResult
            {
                Message = "No attachments found in the email."
            };

        // Resolve before creating anything: a refused destination must not leave a
        // directory behind outside the allowlist.
        outputDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputDir,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputDir");

        try
        {
            Directory.CreateDirectory(outputDir);
        }
        catch (Exception ex)
        {
            throw EmailErrorTranslator.TranslateOutputFailure(ex);
        }

        var extractedFiles = new List<string>();
        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];

        // Two attachments may carry the same name, which is normal in a forwarded mail. Writing
        // both to the same path silently left only the last one, and the reported count still
        // claimed every attachment had been extracted. The resolver gives each a distinct name.
        var collisionResolver = new OleCollisionResolver();

        // One message can carry thousands of parts, and each is written straight to disk. The
        // count is known up front; the size only becomes known as the parts are written, so both
        // are bounded (R2-R02).
        RenderBudget.EnsureOutputCount(message.Attachments.Count, "attachment files");

        // Staged as a batch and published once, so a message whose later attachment passes the
        // limit leaves the destination directory as it found it (R5-R01). The resolver keeps its
        // own set of reserved names, so two attachments sharing a name still get distinct ones
        // even though staging never creates the destination.
        using var batch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "extracted attachments",
            context.Recovery, allowedBasePaths);

        foreach (var attachment in message.Attachments)
        {
            var fileName = SecurityHelper.SanitizeFileName(attachment.Name);
            var outputPath = collisionResolver.Reserve(outputDir, fileName);
            // H42: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
                allowedBasePaths, nameof(outputPath));
            try
            {
                // Measured after the write, the file that passed the limit was already on disk
                // (R3-R07). The bound now applies as the bytes are produced, and the destination
                // only appears once the whole request has succeeded.
                batch.Stage(outputPath, attachment.Save);
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw EmailErrorTranslator.TranslateOutputFailure(ex, fileName);
            }

            extractedFiles.Add(Path.GetFileName(outputPath));
        }

        batch.Publish();

        return new SuccessResult
        {
            Message = $"Extracted {extractedFiles.Count} attachment(s) to '{outputDir}': " +
                      string.Join(", ", extractedFiles) + "."
        };
    }
}
