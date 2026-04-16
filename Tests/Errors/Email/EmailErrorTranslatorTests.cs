using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.Email;

namespace AsposeMcpServer.Tests.Errors.Email;

/// <summary>
///     Unit tests for <see cref="EmailErrorTranslator" />. Verifies that Aspose.Email and BCL
///     exceptions map to the expected sanitized BCL exception types and that no raw
///     inner-exception text (including file paths) leaks through (charter §5 red-line F-10).
/// </summary>
/// <remarks>
///     Aspose.Email 23.10.0 does not expose a dedicated password exception type with no
///     public constructor, so the password path is covered via the fallback
///     <see cref="UnauthorizedAccessException" /> branch (the same branch that the
///     <see cref="EmailErrorTranslator" /> routes to for BCL access-denied failures).
/// </remarks>
public class EmailErrorTranslatorTests
{
    // ─── Path-in-message → no leakage ───────────────────────────────────────────

    [Fact]
    public void Translate_GenericExceptionWithPath_DoesNotContainPath()
    {
        var ex = new Exception("error reading /etc/secret/inbox.eml");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/etc/secret", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("inbox.eml", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_GenericExceptionWithWindowsPath_DoesNotContainPath()
    {
        var ex = new Exception(@"C:\Users\admin\Mail\private.msg not found");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.DoesNotContain("admin", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private.msg", result.Message, StringComparison.Ordinal);
    }

    // ─── Fallback mapping ────────────────────────────────────────────────────────

    [Fact]
    public void Translate_UnknownException_ReturnsInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_UnknownException_MessageIsProcessingFailedSentinel()
    {
        var ex = new InvalidOperationException("boom");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    // ─── BCL exception mappings ──────────────────────────────────────────────────

    [Fact]
    public void Translate_UnauthorizedAccess_ReturnsUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_MessageIsOutputDirectoryNotWritableSentinel()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.OutputDirectoryNotWritable(), result.Message);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_DoesNotLeakPath()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("permission denied", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_DirectoryNotFound_ReturnsUnauthorizedAccessException()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_DirectoryNotFound_DoesNotLeakPath()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/hidden/output", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_WithContextBasename_IncludesBasename()
    {
        var ex = new UnauthorizedAccessException("permission denied");
        var result = EmailErrorTranslator.Translate(ex, "inbox.eml");

        Assert.Contains("inbox.eml", result.Message, StringComparison.Ordinal);
    }

    // ─── No inner-exception attached ─────────────────────────────────────────────

    [Fact]
    public void Translate_Generic_NeverAttachesInnerException()
    {
        var ex = new Exception("boom");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_NeverAttachesInnerException()
    {
        var ex = new UnauthorizedAccessException("perm denied /secret/path");
        var result = EmailErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }
}
