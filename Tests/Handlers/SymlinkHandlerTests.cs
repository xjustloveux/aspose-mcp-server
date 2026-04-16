using System.Reflection;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.Attachment;
using AsposeMcpServer.Handlers.Email.Contact;
using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Handlers.Pdf.Attachment;
using AsposeMcpServer.Handlers.PowerPoint.Background;
using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Tests.Infrastructure;
using BackgroundType = Aspose.Slides.BackgroundType;

namespace AsposeMcpServer.Tests.Handlers;

/// <summary>
///     Regression guard for the symlink handler guard.
///     Each test verifies that a handler (or DocumentConverter) rejects a user-supplied
///     path that is a symlink pointing outside the configured allowlist.
///     Tests are organized into four groups:
///     1. Representative handler integration tests — one per major feature area
///     (Word, Excel, PowerPoint, PDF, Email) exercising the homogeneous
///     <c>ResolveAndEnsureWithinAllowlist</c> insertion pattern.
///     2. B-1/B-2/B-3 specific tests — the three HIGH-severity read-sink blockers
///     (<c>AddPdfAttachmentHandler</c>, <c>SetPhotoEmailContactHandler</c>,
///     <c>SetPptBackgroundHandler</c> / <c>AddImagePptWatermarkHandler</c> /
///     <c>SetBackgroundExcelViewHandler</c>).
///     3. DocumentConverter AllowedBasePaths threading test.
///     4. Negative regression test — a non-symlink path inside the allowlist succeeds.
///     Platform gating: symlink creation requires elevated privileges on Windows (Developer
///     Mode or administrator).  Tests that create symlinks use <c>[SkippableFact]</c> and
///     call <see cref="Skip.IfNot" /> against <see cref="SymlinksAvailable" />.
///     The negative regression test uses plain <c>[Fact]</c> and runs on all platforms.
/// </summary>
public class SymlinkHandlerTests : TestBase
{
    /// <summary>Whether the current OS / privilege level supports symlink creation.</summary>
    private static readonly bool SymlinksAvailable;

    static SymlinkHandlerTests()
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
    ///     Creates an <see cref="OperationContext{TContext}" /> for a handler that does not
    ///     require a real Aspose document object (e.g. Email handlers whose context Document
    ///     is <c>object</c>).
    /// </summary>
    private static OperationContext<object> BuildObjectContext(ServerConfig serverConfig)
    {
        return new OperationContext<object>
        {
            Document = new object(),
            ServerConfig = serverConfig
        };
    }

    /// <summary>Creates a minimal valid VCF contact file at <paramref name="path" />.</summary>
    private static void WriteMinimalVcf(string path)
    {
        // Aspose.Email.Mapi.MapiContact.Save requires a valid VCard structure.
        // We write a raw VCard string that is sufficient for the handler to load.
        File.WriteAllText(path,
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Test User\r\nN:User;Test;;;\r\nEND:VCARD\r\n");
    }

    /// <summary>
    ///     Creates a minimal valid BMP image file at <paramref name="path" />.
    ///     The BMP header is the minimum needed so Aspose can parse the byte array as an image.
    /// </summary>
    private static void WriteMinimalBmp(string path)
    {
        const int width = 4;
        const int height = 4;
        const int rowSize = (width * 24 + 31) / 32 * 4; // row stride padded to 4 bytes
        const int pixelDataSize = rowSize * height;
        const int fileSize = 54 + pixelDataSize;
        var bmp = new byte[fileSize];
        // BM signature
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        // File size (LE)
        bmp[2] = fileSize & 0xFF;
        bmp[3] = 0; // high byte of LE file size — always 0 for this 102-byte fixture
        // Pixel data offset = 54
        bmp[10] = 54;
        // BITMAPINFOHEADER size = 40
        bmp[14] = 40;
        // Width
        bmp[18] = width;
        // Height
        bmp[22] = height;
        // Color planes = 1
        bmp[26] = 1;
        // Bits per pixel = 24
        bmp[28] = 24;
        // Pixel data: fill with red pixels
        for (var i = 54; i < bmp.Length; i += 3)
            if (i + 2 < bmp.Length)
            {
                bmp[i] = 0x00; // Blue
                bmp[i + 1] = 0x00; // Green
                bmp[i + 2] = 0xFF; // Red
            }

        File.WriteAllBytes(path, bmp);
    }

    // =====================================================================
    // Group 1 — Representative handler integration tests (one per feature area)
    // =====================================================================

    // -----------------------------------------------------------------
    // 1a. Word — CreateWordDocumentHandler (.Save sink, H7)
    //     Symlink outputPath escapes allowlist → ArgumentException before write
    // -----------------------------------------------------------------

    /// <summary>
    ///     <c>CreateWordDocumentHandler</c> must throw <see cref="ArgumentException" /> when
    ///     the supplied <c>outputPath</c> is a symlink whose resolved target lies outside the
    ///     configured allowlist.  No file must be written to the symlink target.
    /// </summary>
    [SkippableFact]
    public void CreateWordDocumentHandler_SymlinkedOutputPathOutsideAllowlist_ThrowsAndDoesNotWrite()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot(); // allowlisted dir
        using var outside = SymlinkFixture.AllowlistedTempRoot(); // outside dir

        // The real target lives outside the allowlist.
        var realTarget = Path.Combine(outside.Root, "secret.docx");
        // The symlink lives inside the allowlisted directory.
        var symlink = Path.Combine(inside.Root, "output.docx");
        SymlinkFixture.TryCreateFileSymlink(symlink, realTarget);

        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Document>
        {
            Document = new Document(),
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("outputPath", symlink);

        var handler = new CreateWordDocumentHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(realTarget),
            "Handler must not write to the symlink target that is outside the allowlist");
    }

    // -----------------------------------------------------------------
    // 1b. Excel — SetBackgroundExcelViewHandler (File.ReadAllBytes sink, H50/B-3)
    //     Symlink imagePath escapes allowlist → ArgumentException before read
    //
    // NOTE: This handler is also covered in Group 2 (B-3) as a HIGH blocker.
    //       The test here focuses on the representative pattern (read sink on
    //       Excel handler); the B-3 test below covers all three B-3 handlers.
    // -----------------------------------------------------------------

    /// <summary>
    ///     <c>SetBackgroundExcelViewHandler</c> must throw <see cref="ArgumentException" /> when
    ///     the supplied <c>imagePath</c> is a symlink pointing outside the allowlist.
    ///     The worksheet's <c>BackgroundImage</c> must remain unset.
    /// </summary>
    [SkippableFact]
    public void SetBackgroundExcelViewHandler_SymlinkedImagePathOutsideAllowlist_ThrowsAndDoesNotRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real image lives outside.
        var realImage = Path.Combine(outside.Root, "secret_image.bmp");
        WriteMinimalBmp(realImage);

        var symlink = Path.Combine(inside.Root, "bg_image.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlink, realImage);

        using var workbook = new Workbook();
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Workbook>
        {
            Document = workbook,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("imagePath", symlink);

        var handler = new SetBackgroundExcelViewHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        // BackgroundImage must not be set when the image path symlink escapes the allowlist.
        Assert.Null(workbook.Worksheets[0].BackgroundImage);
    }

    // -----------------------------------------------------------------
    // 1c. PowerPoint — SetPptBackgroundHandler (File.ReadAllBytes read sink, B-3)
    //     Included in Group 2 (B-3) below; listed here as the representative PPT test.
    // -----------------------------------------------------------------

    // -----------------------------------------------------------------
    // 1d. PDF — ExtractEmailAttachmentHandler used as "Email" representative below;
    //     for PDF representative we use AddPdfAttachmentHandler (B-1) in Group 2.
    //     Additional PDF representative: a generic write-sink via CreateWordDocumentHandler
    //     was covered in 1a.  PDF attachment covered fully in B-1.
    // -----------------------------------------------------------------

    // -----------------------------------------------------------------
    // 1e. Email — ExtractEmailAttachmentHandler (File.WriteAllBytes sink via attachment.Save, H42)
    //     Symlink outputDir component escapes allowlist → ArgumentException before write
    // -----------------------------------------------------------------

    /// <summary>
    ///     <c>ExtractEmailAttachmentHandler</c> must throw <see cref="ArgumentException" /> when
    ///     the resolved output path (inside a symlinked output directory) escapes the allowlist.
    ///     No attachment file must be written outside the allowlist.
    /// </summary>
    [SkippableFact]
    public void ExtractEmailAttachmentHandler_SymlinkedOutputDirOutsideAllowlist_ThrowsAndDoesNotWrite()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Create a real directory outside the allowlist.
        var realOutputDir = Path.Combine(outside.Root, "extraction_target");
        Directory.CreateDirectory(realOutputDir);

        // Plant a symlink inside the allowlisted dir that points to the outside dir.
        var symlinkDir = Path.Combine(inside.Root, "output_dir");
        SymlinkFixture.CreateDirSymlink(symlinkDir, realOutputDir);

        // Build a minimal EML file with one attachment.
        var emlPath = Path.Combine(inside.Root, "test.eml");
        File.WriteAllText(emlPath,
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=\"bound1\"\r\n" +
            "Subject: Test\r\n" +
            "\r\n" +
            "--bound1\r\n" +
            "Content-Type: text/plain\r\n" +
            "\r\nBody text\r\n" +
            "--bound1\r\n" +
            "Content-Type: application/octet-stream; name=\"secret.txt\"\r\n" +
            "Content-Disposition: attachment; filename=\"secret.txt\"\r\n" +
            "Content-Transfer-Encoding: base64\r\n" +
            "\r\n" +
            "U0VDUkVU\r\n" + // base64("SECRET")
            "--bound1--\r\n");

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", emlPath);
        parameters.Set("outputDir", symlinkDir);
        parameters.Set("index", 0);

        var handler = new ExtractEmailAttachmentHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));

        // No file should have been written to the outside directory.
        Assert.Empty(Directory.GetFiles(realOutputDir));
    }

    // =====================================================================
    // Group 2 — B-1/B-2/B-3 HIGH-severity read-sink tests (dedicated per handler)
    // =====================================================================

    // -----------------------------------------------------------------
    // B-1: AddPdfAttachmentHandler — symlink attachmentPath outside allowlist
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-1: <c>AddPdfAttachmentHandler</c> reads the attachment via <c>File.ReadAllBytes</c>
    ///     on the resolved path.  When <c>attachmentPath</c> is a symlink pointing outside the
    ///     allowlist, the handler must throw <see cref="ArgumentException" /> before any bytes
    ///     are read.  The PDF's <c>EmbeddedFiles</c> collection must remain empty.
    /// </summary>
    [SkippableFact]
    public void AddPdfAttachmentHandler_SymlinkedAttachmentPathOutsideAllowlist_ThrowsAndDoesNotRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real secret file outside the allowlist.
        var secretFile = Path.Combine(outside.Root, "secret_data.bin");
        File.WriteAllBytes(secretFile, "SECRET"u8.ToArray());

        // Symlink inside allowlist → resolves outside.
        var symlink = Path.Combine(inside.Root, "attachment.bin");
        SymlinkFixture.TryCreateFileSymlink(symlink, secretFile);

        using var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Aspose.Pdf.Document>
        {
            Document = pdfDoc,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("attachmentPath", symlink);
        parameters.Set("attachmentName", "attachment.bin");

        var handler = new AddPdfAttachmentHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.Empty(pdfDoc.EmbeddedFiles);
    }

    // -----------------------------------------------------------------
    // B-2: SetPhotoEmailContactHandler — symlink photoPath outside allowlist
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-2: <c>SetPhotoEmailContactHandler</c> reads the photo via <c>File.ReadAllBytes</c>.
    ///     When <c>photoPath</c> is a symlink whose resolved target is outside the allowlist,
    ///     the handler must throw <see cref="ArgumentException" /> before any bytes are read.
    ///     The output contact file must not be written.
    /// </summary>
    [SkippableFact]
    public void SetPhotoEmailContactHandler_SymlinkedPhotoPathOutsideAllowlist_ThrowsAndDoesNotWrite()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Real photo outside the allowlist.
        var realPhoto = Path.Combine(outside.Root, "secret_photo.bmp");
        WriteMinimalBmp(realPhoto);

        // Symlink inside → resolves outside.
        var symlinkPhoto = Path.Combine(inside.Root, "photo_link.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlinkPhoto, realPhoto);

        // Input VCF lives inside the allowlist.
        var inputVcf = Path.Combine(inside.Root, "contact_in.vcf");
        WriteMinimalVcf(inputVcf);

        var outputVcf = Path.Combine(inside.Root, "contact_out.vcf");

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", inputVcf);
        parameters.Set("outputPath", outputVcf);
        parameters.Set("photoPath", symlinkPhoto);

        var handler = new SetPhotoEmailContactHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(outputVcf),
            "Output contact file must not be written when photoPath symlink escapes the allowlist");
    }

    // -----------------------------------------------------------------
    // B-2 variant: symlink outputPath outside allowlist (write sink)
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-2 write-sink variant: <c>SetPhotoEmailContactHandler</c> calls
    ///     <c>contact.Save(resolvedOutputPath, ...)</c>.  When <c>outputPath</c> is itself a
    ///     symlink pointing outside the allowlist, the handler must throw before writing.
    /// </summary>
    [SkippableFact]
    public void SetPhotoEmailContactHandler_SymlinkedOutputPathOutsideAllowlist_ThrowsAndDoesNotWrite()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        // Legitimate photo inside allowlist.
        var photoPath = Path.Combine(inside.Root, "legit_photo.bmp");
        WriteMinimalBmp(photoPath);

        // Input VCF inside allowlist.
        var inputVcf = Path.Combine(inside.Root, "contact_in2.vcf");
        WriteMinimalVcf(inputVcf);

        // Output path is a symlink pointing outside the allowlist.
        var secretOutput = Path.Combine(outside.Root, "secret_contact.vcf");
        var symlinkOutput = Path.Combine(inside.Root, "output_link.vcf");
        SymlinkFixture.TryCreateFileSymlink(symlinkOutput, secretOutput);

        var serverConfig = BuildServerConfig(inside.Root);
        var context = BuildObjectContext(serverConfig);
        var parameters = new OperationParameters();
        parameters.Set("path", inputVcf);
        parameters.Set("outputPath", symlinkOutput);
        parameters.Set("photoPath", photoPath);

        var handler = new SetPhotoEmailContactHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.False(File.Exists(secretOutput),
            "Handler must not write through a symlinked outputPath that escapes the allowlist");
    }

    // -----------------------------------------------------------------
    // B-3a: SetPptBackgroundHandler — symlink imagePath outside allowlist
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-3a: <c>SetPptBackgroundHandler</c> reads the image via <c>File.ReadAllBytes</c>.
    ///     When <c>imagePath</c> is a symlink pointing outside the allowlist, the handler must
    ///     throw <see cref="ArgumentException" /> before any bytes are read.
    /// </summary>
    [SkippableFact]
    public void SetPptBackgroundHandler_SymlinkedImagePathOutsideAllowlist_ThrowsAndDoesNotRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var realImage = Path.Combine(outside.Root, "secret_bg.bmp");
        WriteMinimalBmp(realImage);

        var symlinkImage = Path.Combine(inside.Root, "bg_link.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlinkImage, realImage);

        using var pres = new Presentation();
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Presentation>
        {
            Document = pres,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("imagePath", symlinkImage);

        var handler = new SetPptBackgroundHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        // Slide background must not have been changed to a picture fill.
        Assert.NotEqual(BackgroundType.OwnBackground, pres.Slides[0].Background.Type);
    }

    // -----------------------------------------------------------------
    // B-3b: AddImagePptWatermarkHandler — symlink imagePath outside allowlist
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-3b: <c>AddImagePptWatermarkHandler</c> reads the watermark image via
    ///     <c>File.ReadAllBytes</c>.  When <c>imagePath</c> is a symlink outside the
    ///     allowlist, the handler must throw before any bytes are read and no shapes
    ///     must be added to the presentation.
    /// </summary>
    [SkippableFact]
    public void AddImagePptWatermarkHandler_SymlinkedImagePathOutsideAllowlist_ThrowsAndDoesNotRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var realImage = Path.Combine(outside.Root, "secret_watermark.bmp");
        WriteMinimalBmp(realImage);

        var symlinkImage = Path.Combine(inside.Root, "watermark_link.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlinkImage, realImage);

        using var pres = new Presentation();
        var shapeCountBefore = pres.Slides[0].Shapes.Count;
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Presentation>
        {
            Document = pres,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("imagePath", symlinkImage);

        var handler = new AddImagePptWatermarkHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.Equal(shapeCountBefore, pres.Slides[0].Shapes.Count);
    }

    // -----------------------------------------------------------------
    // B-3c: SetBackgroundExcelViewHandler — symlink imagePath outside allowlist
    //       (already covered in Group 1 as 1b, repeated here as explicit B-3 entry)
    // -----------------------------------------------------------------

    /// <summary>
    ///     B-3c: <c>SetBackgroundExcelViewHandler</c> reads the background image via
    ///     <c>File.ReadAllBytes</c>.  When <c>imagePath</c> is a symlink outside the
    ///     allowlist, the handler must throw <see cref="ArgumentException" /> before any
    ///     bytes are read and the worksheet <c>BackgroundImage</c> must remain null.
    /// </summary>
    [SkippableFact]
    public void SetBackgroundExcelViewHandler_SymlinkedImagePath_ThrowsBeforeRead()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var realImage = Path.Combine(outside.Root, "secret_excel_bg.bmp");
        WriteMinimalBmp(realImage);

        var symlink = Path.Combine(inside.Root, "excel_bg_link.bmp");
        SymlinkFixture.TryCreateFileSymlink(symlink, realImage);

        using var workbook = new Workbook();
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Workbook>
        {
            Document = workbook,
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("imagePath", symlink);

        var handler = new SetBackgroundExcelViewHandler();

        Assert.Throws<ArgumentException>(() => handler.Execute(context, parameters));
        Assert.Null(workbook.Worksheets[0].BackgroundImage);
    }

    // =====================================================================
    // Group 3 — DocumentConverter AllowedBasePaths threading test
    // =====================================================================

    /// <summary>
    ///     <c>DocumentConverter.ConvertWordDocument</c> uses a private <c>ResolveOutputPath</c>
    ///     helper that delegates to <c>SecurityHelper.ResolveAndEnsureWithinAllowlist</c>.
    ///     When <c>ConversionOptions.AllowedBasePaths</c> is populated and the output path is a
    ///     symlink pointing outside the allowlist, conversion must throw
    ///     <see cref="ArgumentException" /> before any file is written.
    /// </summary>
    [SkippableFact]
    public void DocumentConverter_ConvertWordDocument_SymlinkedOutputPathOutsideAllowlist_Throws()
    {
        Skip.IfNot(SymlinksAvailable, "Symlink creation not available on this platform/privilege level");

        using var inside = SymlinkFixture.AllowlistedTempRoot();
        using var outside = SymlinkFixture.AllowlistedTempRoot();

        var realOutput = Path.Combine(outside.Root, "secret_output.docx");
        var symlinkOutput = Path.Combine(inside.Root, "output_link.docx");
        SymlinkFixture.TryCreateFileSymlink(symlinkOutput, realOutput);

        var doc = new Document();
        var options = new ConversionOptions
        {
            AllowedBasePaths = [inside.Root]
        };

        Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertWordDocument(doc, symlinkOutput, "docx", options: options));
        Assert.False(File.Exists(realOutput),
            "ConvertWordDocument must not write through a symlink that escapes the allowlist");
    }

    // =====================================================================
    // Group 4 — Negative regression test (non-symlink inside allowlist succeeds)
    // =====================================================================

    /// <summary>
    ///     Regression guard: a handler with a regular (non-symlink) output path that lies
    ///     inside the allowlist must not be affected by the symlink checks.
    ///     This test uses <c>CreateWordDocumentHandler</c> and runs on all platforms
    ///     (no symlink creation required).
    /// </summary>
    [Fact]
    public void CreateWordDocumentHandler_LegitimateNonSymlinkPath_Succeeds()
    {
        using var inside = SymlinkFixture.AllowlistedTempRoot();

        var outputPath = Path.Combine(inside.Root, "legit_output.docx");
        var serverConfig = BuildServerConfig(inside.Root);
        var context = new OperationContext<Document>
        {
            Document = new Document(),
            ServerConfig = serverConfig
        };
        var parameters = new OperationParameters();
        parameters.Set("outputPath", outputPath);

        var handler = new CreateWordDocumentHandler();

        // Must not throw; file must be written.
        var result = handler.Execute(context, parameters);

        Assert.NotNull(result);
        Assert.True(File.Exists(outputPath),
            "Handler must write the output file when the path is inside the allowlist");
    }
}
