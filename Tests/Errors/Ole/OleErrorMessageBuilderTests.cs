using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.Ole;

namespace AsposeMcpServer.Tests.Errors.Ole;

/// <summary>
///     Unit tests for <see cref="OleErrorMessageBuilder" />. Verifies every template
///     applies <see cref="AsposeMcpServer.Helpers.Ole.OleSanitizerHelper.SanitizeForLog" />
///     to attacker-reachable fields and never leaks full paths, stack traces, or CRLF
///     content (F-4, F-10).
/// </summary>
public class OleErrorMessageBuilderTests
{
    [Fact]
    public void UnknownOperation_StripsInjection()
    {
        var msg = OleErrorMessageBuilder.UnknownOperation("bogus\r\nFAKE");
        Assert.DoesNotContain('\r', msg);
        Assert.DoesNotContain('\n', msg);
    }

    [Fact]
    public void InvalidPath_UsesOnlyBasename()
    {
        var msg = OleErrorMessageBuilder.InvalidPath("/etc/passwd/../secret/report.xlsx");
        Assert.DoesNotContain("/etc", msg);
        Assert.DoesNotContain("passwd", msg);
        Assert.Contains("report.xlsx", msg);
    }

    [Fact]
    public void InvalidPath_NullReturnsFallbackSentinel()
    {
        var msg = OleErrorMessageBuilder.InvalidPath(null);
        Assert.Contains("invalid", msg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void IndexOutOfRange_IncludesNumericsOnly()
    {
        var msg = OleErrorMessageBuilder.IndexOutOfRange(7, 3);
        Assert.Contains("7", msg);
        Assert.Contains("3", msg);
    }

    [Fact]
    public void OutputDirectoryNotWritable_UsesOnlyBasename()
    {
        var msg = OleErrorMessageBuilder.OutputDirectoryNotWritable("/secret/chain/dir");
        Assert.DoesNotContain("/secret/chain", msg);
    }

    [Fact]
    public void SaveFailed_UsesOnlyBasename()
    {
        var msg = OleErrorMessageBuilder.SaveFailed("/etc/passwd/../report.xlsx");
        Assert.Contains("report.xlsx", msg);
        Assert.DoesNotContain("/etc", msg);
    }

    [Fact]
    public void InvalidPassword_IsFixedSentinel()
    {
        Assert.Equal(OleErrorMessageBuilder.InvalidPasswordSentinel, OleErrorMessageBuilder.InvalidPassword());
    }

    [Fact]
    public void LinkedCannotExtract_IsFixedSentinel()
    {
        Assert.Equal(OleErrorMessageBuilder.LinkedCannotExtractSentinel,
            OleErrorMessageBuilder.LinkedCannotExtract());
    }

    [Fact]
    public void UnsupportedLegacyFormat_Sanitizes()
    {
        var msg = OleErrorMessageBuilder.UnsupportedLegacyFormat(".doc\r\nFAKE");
        Assert.DoesNotContain('\r', msg);
        Assert.DoesNotContain('\n', msg);
    }

    // ─── Regression: OleErrorMessageBuilder.InvalidPassword() delegates to
    //     ErrorMessageBuilder after Phase A refactor. Verify byte-identical output.

    [Fact]
    public void InvalidPassword_DelegatesToUnifiedBuilder_SameValue()
    {
        // OleErrorMessageBuilder.InvalidPassword() now delegates to
        // ErrorMessageBuilder.InvalidPassword(). The returned string must be
        // byte-identical to the unified sentinel so callers of either class
        // observe no behavioral change.
        Assert.Equal(
            ErrorMessageBuilder.InvalidPassword(),
            OleErrorMessageBuilder.InvalidPassword());
    }

    [Fact]
    public void InvalidPassword_ConstantSentinelMatchesUnifiedBuilder()
    {
        // The compile-time constant OleErrorMessageBuilder.InvalidPasswordSentinel
        // is the source of truth for OLE callers; it must equal the unified builder
        // value so that any future change to the sentinel is detected immediately.
        Assert.Equal(
            OleErrorMessageBuilder.InvalidPasswordSentinel,
            ErrorMessageBuilder.InvalidPassword());
    }
}
