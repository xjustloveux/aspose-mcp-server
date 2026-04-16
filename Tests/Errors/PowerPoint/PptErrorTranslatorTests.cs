using System.Reflection;
using Aspose.Slides;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.PowerPoint;

namespace AsposeMcpServer.Tests.Errors.PowerPoint;

/// <summary>
///     Unit tests for <see cref="PptErrorTranslator" />. Verifies that Aspose.Slides and BCL
///     exceptions map to the expected sanitized BCL exception types and that no raw
///     inner-exception text (including file paths) leaks through (charter §5 red-line F-10).
/// </summary>
/// <remarks>
///     <see cref="InvalidPasswordException" /> has no public constructor in Aspose.Slides
///     23.10.0, so it is constructed via private-constructor reflection — analogous to the
///     pattern in <c>CellsErrorTranslatorTests</c>.
/// </remarks>
public class PptErrorTranslatorTests
{
    // ─── factory ──────────────────────────────────────────────────────────────────

    /// <summary>
    ///     Creates an <see cref="InvalidPasswordException" /> via reflection.
    ///     Aspose.Slides 23.10.0 exposes no public constructor for this type.
    /// </summary>
    private static InvalidPasswordException MakeInvalidPasswordException(string message = "wrong pw")
    {
        var type = typeof(InvalidPasswordException);

        // Try string ctor first.
        var stringCtor = type.GetConstructor(
            BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
            null,
            [typeof(string)],
            null);

        if (stringCtor != null)
            return (InvalidPasswordException)stringCtor.Invoke([message]);

        // Parameterless fallback.
        var defaultCtor = type.GetConstructor(
                              BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
                              null,
                              Type.EmptyTypes,
                              null)
                          ?? throw new InvalidOperationException(
                              "Aspose.Slides.InvalidPasswordException: no usable ctor found — Aspose version changed?");

        return (InvalidPasswordException)defaultCtor.Invoke(null);
    }

    // ─── InvalidPasswordException mapping ───────────────────────────────────────

    [Fact]
    public void Translate_InvalidPasswordException_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeInvalidPasswordException();
        var result = PptErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_InvalidPasswordException_MessageIsInvalidPasswordSentinel()
    {
        var ex = MakeInvalidPasswordException("secret internal detail");
        var result = PptErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), result.Message);
    }

    [Fact]
    public void Translate_InvalidPasswordException_DoesNotLeakInnerMessage()
    {
        var ex = MakeInvalidPasswordException("secret internal detail");
        var result = PptErrorTranslator.Translate(ex);

        Assert.DoesNotContain("secret internal detail", result.Message, StringComparison.Ordinal);
    }

    // ─── Path-in-message → no leakage ───────────────────────────────────────────

    [Fact]
    public void Translate_GenericExceptionWithPath_DoesNotContainPath()
    {
        var ex = new Exception("error reading /etc/secret/presentation.pptx");
        var result = PptErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/etc/secret", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("presentation.pptx", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_GenericExceptionWithWindowsPath_DoesNotContainPath()
    {
        var ex = new Exception(@"C:\Users\admin\Documents\private.pptx not found");
        var result = PptErrorTranslator.Translate(ex);

        Assert.DoesNotContain("admin", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private.pptx", result.Message, StringComparison.Ordinal);
    }

    // ─── Fallback mapping ────────────────────────────────────────────────────────

    [Fact]
    public void Translate_UnknownException_ReturnsInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = PptErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_UnknownException_MessageIsProcessingFailedSentinel()
    {
        var ex = new InvalidOperationException("boom");
        var result = PptErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    // ─── BCL exception mappings ──────────────────────────────────────────────────

    [Fact]
    public void Translate_UnauthorizedAccess_ReturnsUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = PptErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_DoesNotLeakPath()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = PptErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("permission denied", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_DirectoryNotFound_ReturnsUnauthorizedAccessException()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = PptErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_WithContextBasename_IncludesBasename()
    {
        var ex = new UnauthorizedAccessException("permission denied");
        var result = PptErrorTranslator.Translate(ex, "deck.pptx");

        Assert.Contains("deck.pptx", result.Message, StringComparison.Ordinal);
    }

    // ─── No inner-exception attached ─────────────────────────────────────────────

    [Fact]
    public void Translate_Password_NeverAttachesInnerException()
    {
        var ex = MakeInvalidPasswordException("internal detail");
        var result = PptErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    [Fact]
    public void Translate_Generic_NeverAttachesInnerException()
    {
        var ex = new Exception("boom");
        var result = PptErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }
}
