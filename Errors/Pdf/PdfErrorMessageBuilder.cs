using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Errors.Pdf;

/// <summary>
///     PDF-family–specific fixed error-message sentinels. Shared cross-family sentinels
///     (invalid password, processing failed, etc.) live in
///     <see cref="ErrorMessageBuilder" />; this class holds only sentinels that are unique
///     to the PDF tool surface.
/// </summary>
public static class PdfErrorMessageBuilder
{
    /// <summary>
    ///     Returns the fixed sentinel used when an individual image within a PDF page
    ///     cannot be accessed or decoded. Designed for the response-field pattern where the
    ///     error is embedded in a <c>PdfImageInfo.Error</c> property rather than thrown, so
    ///     raw Aspose exception text never reaches the MCP caller.
    /// </summary>
    /// <returns>
    ///     The fixed string <c>"Image could not be accessed or decoded."</c>.
    /// </returns>
    [SuppressMessage("SonarAnalyzer.CSharp", "S3400",
        Justification = "Intentional method form, paralleling the cross-family ErrorMessageBuilder " +
                        "API shape. Converting to a field would break uniformity with peer sentinels " +
                        "and response-field embedding patterns (PdfImageInfo.Error).")]
    public static string ImageAccessError()
    {
        return "Image could not be accessed or decoded.";
    }
}
