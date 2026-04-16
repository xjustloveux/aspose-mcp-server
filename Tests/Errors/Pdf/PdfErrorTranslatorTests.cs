using System.Reflection;
using Aspose.Pdf;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.Pdf;

namespace AsposeMcpServer.Tests.Errors.Pdf;

/// <summary>
///     Unit tests for <see cref="PdfErrorTranslator" />. Verifies that Aspose.PDF and BCL
///     exceptions map to the expected sanitized BCL exception types and that no raw
///     inner-exception text (including file paths) leaks through (charter §5 red-line F-10).
/// </summary>
/// <remarks>
///     <see cref="Aspose.Pdf.InvalidPasswordException" /> has no public constructor in
///     Aspose.PDF 23.10.0, so instances are created via private-constructor reflection,
///     analogous to the pattern in <c>CellsErrorTranslatorTests</c>.
/// </remarks>
public class PdfErrorTranslatorTests
{
    // ─── factory ──────────────────────────────────────────────────────────────────

    /// <summary>
    ///     Creates an <see cref="Aspose.Pdf.InvalidPasswordException" /> via reflection.
    ///     The exception has no public constructor in Aspose.PDF 23.10.0; a string-only
    ///     private constructor is used as a fallback when the parameterless form is absent.
    /// </summary>
    private static InvalidPasswordException MakeInvalidPasswordException(
        string message = "wrong password")
    {
        var type = typeof(InvalidPasswordException);

        // Try parameterless ctor first (some builds expose it).
        var defaultCtor = type.GetConstructor(
            BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
            null,
            Type.EmptyTypes,
            null);
        if (defaultCtor != null)
            return (InvalidPasswordException)defaultCtor.Invoke(null);

        // Fall back to (string) ctor.
        var stringCtor = type.GetConstructor(
                             BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
                             null,
                             [typeof(string)],
                             null)
                         ?? throw new InvalidOperationException(
                             "Aspose.Pdf.InvalidPasswordException: no usable ctor found — Aspose version changed?");

        return (InvalidPasswordException)stringCtor.Invoke([message]);
    }

    // ─── InvalidPasswordException mapping ───────────────────────────────────────

    [Fact]
    public void Translate_InvalidPasswordException_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeInvalidPasswordException("wrong pw");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_InvalidPasswordException_MessageIsInvalidPasswordSentinel()
    {
        var ex = MakeInvalidPasswordException("secret internal detail");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), result.Message);
    }

    [Fact]
    public void Translate_InvalidPasswordException_DoesNotLeakInnerMessage()
    {
        var ex = MakeInvalidPasswordException("secret internal detail");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.DoesNotContain("secret internal detail", result.Message, StringComparison.Ordinal);
    }

    // ─── Path-in-message → no leakage ───────────────────────────────────────────

    [Fact]
    public void Translate_GenericExceptionWithPath_DoesNotContainPath()
    {
        var ex = new Exception("could not open /etc/secret/document.pdf");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/etc/secret", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("document.pdf", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_GenericExceptionWithWindowsPath_DoesNotContainPath()
    {
        var ex = new Exception(@"C:\Users\admin\Documents\private.pdf not found");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.DoesNotContain("admin", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private.pdf", result.Message, StringComparison.Ordinal);
    }

    // ─── Fallback mapping ────────────────────────────────────────────────────────

    [Fact]
    public void Translate_UnknownException_ReturnsInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_UnknownException_MessageIsProcessingFailedSentinel()
    {
        var ex = new InvalidOperationException("boom");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    // ─── BCL exception mappings ──────────────────────────────────────────────────

    [Fact]
    public void Translate_UnauthorizedAccess_ReturnsUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_DoesNotLeakPath()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("permission denied", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_DirectoryNotFound_ReturnsUnauthorizedAccessException()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_WithContextBasename_IncludesBasename()
    {
        var ex = new UnauthorizedAccessException("permission denied");
        var result = PdfErrorTranslator.Translate(ex, "report.pdf");

        Assert.Contains("report.pdf", result.Message, StringComparison.Ordinal);
    }

    // ─── No inner-exception attached ─────────────────────────────────────────────

    [Fact]
    public void Translate_NeverAttachesInnerException()
    {
        var ex = MakeInvalidPasswordException("internal detail");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    [Fact]
    public void Translate_Generic_NeverAttachesInnerException()
    {
        var ex = new Exception("boom");
        var result = PdfErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    // ─── TranslateImageAccessError ───────────────────────────────────────────────

    [Fact]
    public void TranslateImageAccessError_InvalidPasswordException_ReturnsInvalidPasswordSentinel()
    {
        var ex = MakeInvalidPasswordException("pw");
        var msg = PdfErrorTranslator.TranslateImageAccessError(ex);

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), msg);
    }

    [Fact]
    public void TranslateImageAccessError_GenericException_ReturnsImageAccessErrorSentinel()
    {
        var ex = new Exception("decode failure /internal/path");
        var msg = PdfErrorTranslator.TranslateImageAccessError(ex);

        Assert.Equal(PdfErrorMessageBuilder.ImageAccessError(), msg);
    }

    [Fact]
    public void TranslateImageAccessError_GenericException_DoesNotContainExceptionMessage()
    {
        var ex = new Exception("decode failure /internal/path");
        var msg = PdfErrorTranslator.TranslateImageAccessError(ex);

        Assert.DoesNotContain("decode failure", msg, StringComparison.Ordinal);
        Assert.DoesNotContain("/internal/path", msg, StringComparison.Ordinal);
    }
}
