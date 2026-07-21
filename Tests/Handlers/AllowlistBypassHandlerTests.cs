using System.Reflection;
using Aspose.Pdf;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Handlers.Word.DigitalSignature;
using AsposeMcpServer.Tests.Infrastructure;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Handlers;

/// <summary>
///     Regression guard for bug 20260416-handler-allowlist-bypass.
///     Two handler read sinks (<c>AddEmailAttachmentHandler</c> and
///     <c>EditPdfImageHandler</c>) previously had no
///     <see cref="AsposeMcpServer.Helpers.SecurityHelper.ResolveAndEnsureWithinAllowlist" />
///     call.  The fix added those calls; these tests verify the guard is in place.
///     Pattern is identical to <see cref="SymlinkHandlerTests" /> (the symlink handler guard).
///     Platform gating: symlink creation requires elevated privileges on Windows (Developer
///     Mode or administrator). Tests that create symlinks use <c>[SkippableFact]</c> and
///     call <see cref="Skip.IfNot" /> against <see cref="SymlinksAvailable" />.
///     The negative regression test uses plain <c>[Fact]</c> and runs on all platforms.
/// </summary>
public class AllowlistBypassHandlerTests : TestBase
{
    /// <summary>Whether the current OS / privilege level supports symlink creation.</summary>
    private static readonly bool SymlinksAvailable;

    static AllowlistBypassHandlerTests()
    {
        using var probe = SymlinkFixture.AllowlistedTempRoot();
        var probeLink = Path.Combine(probe.Root, "probe_link");
        var probeTarget = Path.Combine(probe.Root, "probe_target.txt");
        File.WriteAllText(probeTarget, "probe");
        SymlinksAvailable = SymlinkFixture.TryCreateFileSymlink(probeLink, probeTarget);
    }

    // ------------------------------------------------------------------
    // Shared helpers
    // ------------------------------------------------------------------

    /// <summary>
    ///     Builds a <see cref="ServerConfig" /> whose <see cref="ServerConfig.AllowedBasePaths" />
    ///     is set via reflection (the property setter is private).
    /// </summary>
    /// <param name="allowedPaths">Zero or more absolute or relative paths to include.</param>
    /// <returns>A configured <see cref="ServerConfig" /> instance.</returns>
    private static ServerConfig BuildServerConfig(params string[] allowedPaths)
    {
        var cfg = new ServerConfig();
        var prop = typeof(ServerConfig).GetProperty(
            nameof(ServerConfig.AllowedBasePaths),
            BindingFlags.Instance | BindingFlags.Public);
        prop!.SetValue(cfg,
            allowedPaths
                .Select(Path.GetFullPath)
                .ToList()
                .AsReadOnly());
        return cfg;
    }

    /// <summary>
    ///     Creates an <see cref="OperationContext{TContext}" /> with a plain <c>object</c> document
    ///     (used for Email handlers whose generic parameter is <c>object</c>).
    /// </summary>
    /// <param name="serverConfig">The server configuration to attach.</param>
    /// <returns>A minimal operation context.</returns>
    private static OperationContext<object> BuildObjectContext(ServerConfig serverConfig)
    {
        return new OperationContext<object>
        {
            Document = new object(),
            ServerConfig = serverConfig
        };
    }

    /// <summary>
    ///     Creates a minimal valid EML file at <paramref name="path" /> containing one attachment
    ///     so that <c>AddEmailAttachmentHandler</c> can load it without error.
    /// </summary>
    /// <param name="path">Destination path for the EML file.</param>
    private static void WriteMinimalEml(string path)
    {
        File.WriteAllText(path,
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/plain\r\n" +
            "Subject: Test\r\n" +
            "From: test@example.com\r\n" +
            "To: dest@example.com\r\n" +
            "\r\nBody text\r\n");
    }

    /// <summary>
    ///     Creates a minimal valid BMP image at <paramref name="path" /> (4 x 4 pixels, 24-bit).
    /// </summary>
    /// <param name="path">Destination path for the BMP file.</param>
    private static void WriteMinimalBmp(string path)
    {
        const int width = 4;
        const int height = 4;
        const int rowSize = (width * 24 + 31) / 32 * 4;
        const int pixelDataSize = rowSize * height;
        const int fileSize = 54 + pixelDataSize;
        var bmp = new byte[fileSize];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        bmp[2] = fileSize & 0xFF;
        bmp[3] = 0; // high byte of LE file size — always 0 for this 102-byte fixture
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = width;
        bmp[22] = height;
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
            if (i + 2 < bmp.Length)
            {
                bmp[i] = 0x00;
                bmp[i + 1] = 0x00;
                bmp[i + 2] = 0xFF;
            }

        File.WriteAllBytes(path, bmp);
    }

    /// <summary>
    ///     Creates a PDF document that already contains one image on page 1 so that
    ///     <c>EditPdfImageHandler</c> can reach the <c>imagePath</c> allowlist check.
    /// </summary>
    /// <returns>A <see cref="Document" /> with one image on page 1.</returns>
    private static Document CreatePdfDocumentWithImage()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        using var ms = new MemoryStream();
        // Minimal 4×4 BMP bytes written to a stream.
        const int width = 4;
        const int height = 4;
        const int rowSize = (width * 24 + 31) / 32 * 4;
        const int pixelDataSize = rowSize * height;
        const int bmpSize = 54 + pixelDataSize;
        var bmp = new byte[bmpSize];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        bmp[2] = bmpSize & 0xFF;
        bmp[3] = 0; // high byte of LE file size — always 0 for this 102-byte fixture
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = width;
        bmp[22] = height;
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
            if (i + 2 < bmp.Length)
            {
                bmp[i] = 0x00;
                bmp[i + 1] = 0xFF;
                bmp[i + 2] = 0x00;
            }

        ms.Write(bmp, 0, bmp.Length);
        ms.Position = 0;
        page.AddImage(ms, new Rectangle(100, 600, 300, 800));
        return doc;
    }

    // =====================================================================
    // Test 1 — AddEmailAttachmentHandler: symlink attachmentPath escapes allowlist
    // =====================================================================

    /// <summary>
    ///     <c>AddEmailAttachmentHandler</c> must throw <see cref="ArgumentException" /> when
    ///     <c>attachmentPath</c> is a symlink whose resolved target lies outside the configured
    ///     allowlist.  The fix added <c>ResolveAndEnsureWithinAllowlist</c> for this parameter
    ///     (bug 20260416-handler-allowlist-bypass).
    /// </summary>
    [SkippableFact]
    public void AddEmailAttachmentHandler_SymlinkedAttachmentPathOutsideAllowlist_ThrowsBeforeRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real secret file lives outside the allowlist.
        var secretFile = Path.Combine(outside.Root, "secret_attachment.bin");
        File.WriteAllBytes(secretFile, "SECRET"u8.ToArray());

        // Symlink inside allowlist → resolves to the outside secret file.
        var symlinkAttachment = Path.Combine(inside.Root, "attachment_link.bin");
        SymlinkFixture.TryCreateFileSymlink(symlinkAttachment, secretFile);

        // Source email and output path are both inside the allowlist.
        var emailPath = Path.Combine(inside.Root, "source.eml");
        WriteMinimalEml(emailPath);
        var outputPath = Path.Combine(inside.Root, "output.eml");

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", emailPath);
        parameters.Set("outputPath", outputPath);
        parameters.Set("attachmentPath", symlinkAttachment);

        var handler = new AddEmailAttachmentHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        // Output must not have been written — handler aborted before reading the attachment.
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when attachmentPath symlink escapes the allowlist");
    }

    // =====================================================================
    // Test 2 — AddEmailAttachmentHandler: symlink path (source email) escapes allowlist
    // =====================================================================

    /// <summary>
    ///     <c>AddEmailAttachmentHandler</c> must throw <see cref="ArgumentException" /> when
    ///     the source email <c>path</c> parameter is a symlink that resolves outside the allowlist.
    ///     The fix added <c>ResolveAndEnsureWithinAllowlist</c> for both read sinks.
    /// </summary>
    [SkippableFact]
    public void AddEmailAttachmentHandler_SymlinkedSourceEmailPathOutsideAllowlist_ThrowsBeforeRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real source email lives outside the allowlist.
        var secretEmail = Path.Combine(outside.Root, "secret.eml");
        WriteMinimalEml(secretEmail);

        // Symlink inside allowlist → resolves to the outside email file.
        var symlinkEmail = Path.Combine(inside.Root, "email_link.eml");
        SymlinkFixture.TryCreateFileSymlink(symlinkEmail, secretEmail);

        // Attachment and output are inside the allowlist.
        var attachmentPath = Path.Combine(inside.Root, "attachment.bin");
        File.WriteAllBytes(attachmentPath, "ABC"u8.ToArray());
        var outputPath = Path.Combine(inside.Root, "output.eml");

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", symlinkEmail);
        parameters.Set("outputPath", outputPath);
        parameters.Set("attachmentPath", attachmentPath);

        var handler = new AddEmailAttachmentHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when source email path symlink escapes the allowlist");
    }

    // =====================================================================
    // Test 3 — EditPdfImageHandler: symlink imagePath escapes allowlist
    // =====================================================================

    /// <summary>
    ///     <c>EditPdfImageHandler</c> must throw <see cref="ArgumentException" /> when
    ///     <c>imagePath</c> is a symlink whose resolved target lies outside the configured
    ///     allowlist.  The fix added <c>ResolveAndEnsureWithinAllowlist</c> before the
    ///     <c>AddImage</c> / <c>page.AddImage</c> call (bug 20260416-handler-allowlist-bypass).
    /// </summary>
    [SkippableFact]
    public void EditPdfImageHandler_SymlinkedImagePathOutsideAllowlist_ThrowsBeforeAddImage()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real secret image lives outside the allowlist.
        var secretImage = Path.Combine(outside.Root, "secret_image.bmp");
        WriteMinimalBmp(secretImage);

        // Symlink inside allowlist → resolves to the outside secret image.
        var symlinkImage = Path.Combine(inside.Root, "image_link.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlinkImage, secretImage);

        using var pdfDoc = CreatePdfDocumentWithImage();

        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Document>
        {
            Document = pdfDoc,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", 1);
        parameters.Set("imageIndex", 1);
        parameters.Set("imagePath", symlinkImage);

        var handler = new EditPdfImageHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        // The image count on the page should not have changed to a replaced-in image from outside.
        // After the fix throws, the delete already happened but AddImage was never called.
        // What matters is that no secret data was read. We verify the exception was raised.
    }

    // =====================================================================
    // Test 4 (negative regression) — legitimate paths inside allowlist succeed
    // =====================================================================

    /// <summary>
    ///     Regression guard: <c>AddEmailAttachmentHandler</c> must still function normally when
    ///     all paths are non-symlink files that reside inside the configured allowlist.
    ///     Runs on all platforms (no symlink creation required).
    /// </summary>
    [Fact]
    public void AddEmailAttachmentHandler_LegitimatePathsInsideAllowlist_Succeeds()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();

        var emailPath = Path.Combine(inside.Root, "source.eml");
        WriteMinimalEml(emailPath);

        var attachmentPath = Path.Combine(inside.Root, "attach.txt");
        File.WriteAllText(attachmentPath, "hello");

        var outputPath = Path.Combine(inside.Root, "output.eml");

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", emailPath);
        parameters.Set("outputPath", outputPath);
        parameters.Set("attachmentPath", attachmentPath);

        var handler = new AddEmailAttachmentHandler();

        // Must not throw; output file must be written.
        var result = handler.Execute(context, parameters);

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath),
            "Handler must write the output email when all paths are inside the allowlist");
    }

    // =====================================================================
    // PPT file-operation input read sinks (bug: input paths validated for
    // shape only, never resolved against the allowlist)
    // =====================================================================

    /// <summary>
    ///     Creates a minimal one-slide presentation file at <paramref name="path" />.
    /// </summary>
    /// <param name="path">Destination path for the PPTX file.</param>
    private static void WriteMinimalPptx(string path)
    {
        using var pres = new Presentation();
        pres.Save(path, SlidesSaveFormat.Pptx);
    }

    /// <summary>
    ///     Creates an <see cref="OperationContext{TContext}" /> for PPT file-operation handlers.
    ///     The document is unused by these handlers (they open their own inputs).
    /// </summary>
    /// <param name="serverConfig">The server configuration to attach.</param>
    /// <returns>A minimal operation context.</returns>
    private static OperationContext<Presentation> BuildPresentationContext(ServerConfig serverConfig)
    {
        return new OperationContext<Presentation>
        {
            Document = null!,
            ServerConfig = serverConfig
        };
    }

    /// <summary>
    ///     <c>ConvertPresentationHandler</c> must reject an input path outside the configured
    ///     allowlist before opening it (the input is a read sink like the output).
    /// </summary>
    [Fact]
    public void ConvertPresentationHandler_InputPathOutsideAllowlist_ThrowsBeforeRead()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var secretPptx = Path.Combine(outside.Root, "secret.pptx");
        WriteMinimalPptx(secretPptx);
        var outputPath = Path.Combine(inside.Root, "converted.pdf");

        var context = BuildPresentationContext(BuildServerConfig(inside.Root));
        var parameters = new OperationParameters();
        parameters.Set("inputPath", secretPptx);
        parameters.Set("outputPath", outputPath);
        parameters.Set("format", "pdf");

        var handler = new ConvertPresentationHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when the input path is outside the allowlist");
    }

    /// <summary>
    ///     <c>SplitPresentationHandler</c> must reject an input path outside the configured
    ///     allowlist before opening it.
    /// </summary>
    [Fact]
    public void SplitPresentationHandler_InputPathOutsideAllowlist_ThrowsBeforeRead()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var secretPptx = Path.Combine(outside.Root, "secret.pptx");
        WriteMinimalPptx(secretPptx);
        var outputDir = Path.Combine(inside.Root, "split_out");

        var context = BuildPresentationContext(BuildServerConfig(inside.Root));
        var parameters = new OperationParameters();
        parameters.Set("inputPath", secretPptx);
        parameters.Set("outputDirectory", outputDir);

        var handler = new SplitPresentationHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
    }

    /// <summary>
    ///     <c>MergePresentationsHandler</c> must reject input paths outside the configured
    ///     allowlist before opening any of them.
    /// </summary>
    [Fact]
    public void MergePresentationsHandler_InputPathOutsideAllowlist_ThrowsBeforeRead()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var insidePptx = Path.Combine(inside.Root, "a.pptx");
        WriteMinimalPptx(insidePptx);
        var secretPptx = Path.Combine(outside.Root, "secret.pptx");
        WriteMinimalPptx(secretPptx);
        var outputPath = Path.Combine(inside.Root, "merged.pptx");

        var context = BuildPresentationContext(BuildServerConfig(inside.Root));
        var parameters = new OperationParameters();
        parameters.Set("inputPaths", new[] { insidePptx, secretPptx });
        parameters.Set("outputPath", outputPath);

        var handler = new MergePresentationsHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when an input path is outside the allowlist");
    }

    // =====================================================================
    // Word digital-signature file sinks (bug: path/outputPath/certificatePath
    // validated for shape only, never resolved against the allowlist)
    // =====================================================================

    /// <summary>
    ///     <c>SignWordDigitalSignatureHandler</c> must reject a certificate path outside the
    ///     configured allowlist before touching the certificate file.
    /// </summary>
    [Fact]
    public void SignWordDigitalSignatureHandler_CertificatePathOutsideAllowlist_Throws()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var docPath = Path.Combine(inside.Root, "doc.docx");
        new Aspose.Words.Document().Save(docPath);
        var outputPath = Path.Combine(inside.Root, "signed.docx");
        var certPath = Path.Combine(outside.Root, "secret_cert.pfx");
        File.WriteAllBytes(certPath, [0x01, 0x02]);

        var context = new OperationContext<Aspose.Words.Document>
        {
            Document = null!,
            ServerConfig = BuildServerConfig(inside.Root)
        };
        var parameters = new OperationParameters();
        parameters.Set("path", docPath);
        parameters.Set("outputPath", outputPath);
        parameters.Set("certificatePath", certPath);
        parameters.Set("certificatePassword", "pw");

        var handler = new SignWordDigitalSignatureHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when the certificate path is outside the allowlist");
    }

    /// <summary>
    ///     <c>RemoveWordDigitalSignatureHandler</c> must reject a source path outside the
    ///     configured allowlist before reading it.
    /// </summary>
    [Fact]
    public void RemoveWordDigitalSignatureHandler_SourcePathOutsideAllowlist_Throws()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var docPath = Path.Combine(outside.Root, "secret.docx");
        new Aspose.Words.Document().Save(docPath);
        var outputPath = Path.Combine(inside.Root, "unsigned.docx");

        var context = new OperationContext<Aspose.Words.Document>
        {
            Document = null!,
            ServerConfig = BuildServerConfig(inside.Root)
        };
        var parameters = new OperationParameters();
        parameters.Set("path", docPath);
        parameters.Set("outputPath", outputPath);

        var handler = new RemoveWordDigitalSignatureHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputPath),
            "Handler must not write output when the source path is outside the allowlist");
    }

    /// <summary>
    ///     Regression guard: <c>ConvertPresentationHandler</c> must still convert normally when
    ///     input and output are both inside the allowlist.
    /// </summary>
    [Fact]
    public void ConvertPresentationHandler_LegitimatePathsInsideAllowlist_Succeeds()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();

        var inputPptx = Path.Combine(inside.Root, "input.pptx");
        WriteMinimalPptx(inputPptx);
        var outputPath = Path.Combine(inside.Root, "converted.pptx");

        var context = BuildPresentationContext(BuildServerConfig(inside.Root));
        var parameters = new OperationParameters();
        parameters.Set("inputPath", inputPptx);
        parameters.Set("outputPath", outputPath);
        parameters.Set("format", "pptx");

        var handler = new ConvertPresentationHandler();

        var result = handler.Execute(context, parameters);

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath),
            "Handler must write the output when all paths are inside the allowlist");
    }
}
