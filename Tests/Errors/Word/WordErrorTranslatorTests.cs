using System.Reflection;
using Aspose.Words;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.Word;

namespace AsposeMcpServer.Tests.Errors.Word;

/// <summary>
///     Unit tests for <see cref="WordErrorTranslator" />. Verifies that Aspose.Words and BCL
///     exceptions map to the expected sanitized BCL exception types and that no raw
///     inner-exception text (including file paths) leaks through (charter §5 red-line F-10).
/// </summary>
/// <remarks>
///     <see cref="IncorrectPasswordException" /> has no public constructor in Aspose.Words
///     23.10.0, so it is constructed via private-constructor reflection — analogous to the
///     pattern in <c>CellsErrorTranslatorTests</c>.
/// </remarks>
public class WordErrorTranslatorTests
{
    // ─── factory ──────────────────────────────────────────────────────────────────

    /// <summary>
    ///     Creates an <see cref="IncorrectPasswordException" /> via reflection.
    ///     Aspose.Words 23.10.0 exposes no public constructor for this type.
    /// </summary>
    private static IncorrectPasswordException MakeIncorrectPasswordException(string message = "wrong pw")
    {
        var type = typeof(IncorrectPasswordException);

        // Attempt string ctor (most likely present in non-public surface).
        var stringCtor = type.GetConstructor(
            BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
            null,
            [typeof(string)],
            null);

        if (stringCtor != null)
            return (IncorrectPasswordException)stringCtor.Invoke([message]);

        // Parameterless fallback.
        var defaultCtor = type.GetConstructor(
                              BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
                              null,
                              Type.EmptyTypes,
                              null)
                          ?? throw new InvalidOperationException(
                              "IncorrectPasswordException: no usable ctor found — Aspose.Words version changed?");

        return (IncorrectPasswordException)defaultCtor.Invoke(null);
    }

    // ─── IncorrectPasswordException mapping ──────────────────────────────────────

    [Fact]
    public void Translate_IncorrectPasswordException_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeIncorrectPasswordException();
        var result = WordErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_IncorrectPasswordException_MessageIsInvalidPasswordSentinel()
    {
        var ex = MakeIncorrectPasswordException("secret internal detail");
        var result = WordErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), result.Message);
    }

    [Fact]
    public void Translate_IncorrectPasswordException_DoesNotLeakInnerMessage()
    {
        var ex = MakeIncorrectPasswordException("secret internal detail");
        var result = WordErrorTranslator.Translate(ex);

        Assert.DoesNotContain("secret internal detail", result.Message, StringComparison.Ordinal);
    }

    // ─── UnsupportedFileFormatException mapping ──────────────────────────────────

    /// <summary>
    ///     Creates an <see cref="UnsupportedFileFormatException" /> via reflection.
    ///     Aspose.Words 23.10.0 does not expose a public string-only constructor for
    ///     this type; the internal/private form is used.
    /// </summary>
    private static UnsupportedFileFormatException MakeUnsupportedFileFormatException()
    {
        var type = typeof(UnsupportedFileFormatException);

        // Parameterless ctor is present in most Aspose.Words builds.
        var defaultCtor = type.GetConstructor(
            BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
            null,
            Type.EmptyTypes,
            null);

        if (defaultCtor != null)
            return (UnsupportedFileFormatException)defaultCtor.Invoke(null);

        // Fall back to any non-public single-param string ctor.
        var stringCtor = type.GetConstructor(
                             BindingFlags.NonPublic | BindingFlags.Instance,
                             null,
                             [typeof(string)],
                             null)
                         ?? throw new InvalidOperationException(
                             "UnsupportedFileFormatException: no usable ctor found — Aspose.Words version changed?");

        return (UnsupportedFileFormatException)stringCtor.Invoke(["unsupported format"]);
    }

    [Fact]
    public void Translate_UnsupportedFileFormatException_ReturnsNotSupportedException()
    {
        var ex = MakeUnsupportedFileFormatException();
        var result = WordErrorTranslator.Translate(ex);

        Assert.IsType<NotSupportedException>(result);
    }

    [Fact]
    public void Translate_UnsupportedFileFormatException_MessageIsSentinel()
    {
        var ex = MakeUnsupportedFileFormatException();
        var result = WordErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.UnsupportedFormat(), result.Message);
    }

    // ─── Path-in-message → no leakage ───────────────────────────────────────────

    [Fact]
    public void Translate_GenericExceptionWithPath_DoesNotContainPath()
    {
        var ex = new Exception("error reading /etc/secret/document.docx");
        var result = WordErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/etc/secret", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("document.docx", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_GenericExceptionWithWindowsPath_DoesNotContainPath()
    {
        var ex = new Exception(@"C:\Users\admin\Documents\private.docx not found");
        var result = WordErrorTranslator.Translate(ex);

        Assert.DoesNotContain("admin", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private.docx", result.Message, StringComparison.Ordinal);
    }

    // ─── Fallback mapping ────────────────────────────────────────────────────────

    [Fact]
    public void Translate_UnknownException_ReturnsInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = WordErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_UnknownException_MessageIsProcessingFailedSentinel()
    {
        var ex = new InvalidOperationException("boom");
        var result = WordErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    // ─── BCL exception mappings ──────────────────────────────────────────────────

    [Fact]
    public void Translate_UnauthorizedAccess_ReturnsUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = WordErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_DoesNotLeakPath()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = WordErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("permission denied", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_DirectoryNotFound_ReturnsUnauthorizedAccessException()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = WordErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_WithContextBasename_IncludesBasename()
    {
        var ex = new UnauthorizedAccessException("permission denied");
        var result = WordErrorTranslator.Translate(ex, "report.docx");

        Assert.Contains("report.docx", result.Message, StringComparison.Ordinal);
    }

    // ─── No inner-exception attached ─────────────────────────────────────────────

    [Fact]
    public void Translate_Password_NeverAttachesInnerException()
    {
        var ex = MakeIncorrectPasswordException("internal detail");
        var result = WordErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    [Fact]
    public void Translate_Generic_NeverAttachesInnerException()
    {
        var ex = new Exception("boom");
        var result = WordErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }
}
