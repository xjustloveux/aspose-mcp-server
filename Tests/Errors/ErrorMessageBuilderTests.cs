using AsposeMcpServer.Errors;

namespace AsposeMcpServer.Tests.Errors;

/// <summary>
///     Unit tests for <see cref="ErrorMessageBuilder" />. Verifies every sentinel method
///     returns a non-empty fixed string and that no variable content (exception text,
///     stack traces, or file-path separators from internal data) leaks through.
/// </summary>
public class ErrorMessageBuilderTests
{
    // ─── helpers ────────────────────────────────────────────────────────────────

    /// <summary>
    ///     Asserts that the message does not contain patterns that indicate raw exception
    ///     or stack-trace leakage.
    /// </summary>
    private static void AssertNoLeakPatterns(string msg)
    {
        Assert.DoesNotContain("Exception", msg, StringComparison.Ordinal);
        // Stack-frame lines start with "   at " (three spaces then "at "); check that
        // specific prefix so the word "format" or "that" does not trigger a false positive.
        Assert.DoesNotContain("   at ", msg, StringComparison.Ordinal);
        Assert.DoesNotContain("StackTrace", msg, StringComparison.Ordinal);
    }

    // ─── InvalidPassword ────────────────────────────────────────────────────────

    [Fact]
    public void InvalidPassword_ReturnsExpectedSentinel()
    {
        var msg = ErrorMessageBuilder.InvalidPassword();
        Assert.Equal(
            "The source file requires a password, or the supplied password is incorrect.",
            msg);
    }

    [Fact]
    public void InvalidPassword_IsNonEmpty()
    {
        Assert.NotEmpty(ErrorMessageBuilder.InvalidPassword());
    }

    [Fact]
    public void InvalidPassword_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.InvalidPassword());
    }

    // ─── IndexOutOfRange ────────────────────────────────────────────────────────

    [Fact]
    public void IndexOutOfRange_NoParams_ReturnsGenericSentinel()
    {
        var msg = ErrorMessageBuilder.IndexOutOfRange();
        Assert.Equal("The supplied index is out of range.", msg);
    }

    [Fact]
    public void IndexOutOfRange_WithItemName_IncludesItemName()
    {
        var msg = ErrorMessageBuilder.IndexOutOfRange("worksheet");
        Assert.Contains("worksheet", msg, StringComparison.Ordinal);
        Assert.Contains("out of range", msg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void IndexOutOfRange_WithBothParams_IncludesBothTokens()
    {
        var msg = ErrorMessageBuilder.IndexOutOfRange("row", "0–99");
        Assert.Contains("row", msg, StringComparison.Ordinal);
        Assert.Contains("0–99", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void IndexOutOfRange_NullItemName_NullRange_ReturnsGeneric()
    {
        var msg = ErrorMessageBuilder.IndexOutOfRange();
        Assert.Equal("The supplied index is out of range.", msg);
    }

    [Fact]
    public void IndexOutOfRange_EmptyStrings_ReturnsGeneric()
    {
        var msg = ErrorMessageBuilder.IndexOutOfRange(string.Empty, string.Empty);
        Assert.Equal("The supplied index is out of range.", msg);
    }

    [Fact]
    public void IndexOutOfRange_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.IndexOutOfRange("col", "0–255"));
    }

    // ─── OutputDirectoryNotWritable ─────────────────────────────────────────────

    [Fact]
    public void OutputDirectoryNotWritable_NoParam_ReturnsGenericSentinel()
    {
        var msg = ErrorMessageBuilder.OutputDirectoryNotWritable();
        Assert.Equal("The output directory cannot be created or is not writable.", msg);
    }

    [Fact]
    public void OutputDirectoryNotWritable_WithContext_IncludesContext()
    {
        var msg = ErrorMessageBuilder.OutputDirectoryNotWritable("report.xlsx");
        Assert.Contains("report.xlsx", msg, StringComparison.Ordinal);
        Assert.Contains("writable", msg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OutputDirectoryNotWritable_NullParam_ReturnsGeneric()
    {
        var msg = ErrorMessageBuilder.OutputDirectoryNotWritable();
        Assert.Equal("The output directory cannot be created or is not writable.", msg);
    }

    [Fact]
    public void OutputDirectoryNotWritable_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.OutputDirectoryNotWritable("out.xlsx"));
    }

    // ─── OperationFailed ────────────────────────────────────────────────────────

    [Fact]
    public void OperationFailed_DefaultVerb_ReturnsProcessSentinel()
    {
        var msg = ErrorMessageBuilder.OperationFailed();
        Assert.Equal("Failed to process the document.", msg);
    }

    [Fact]
    public void OperationFailed_CustomVerb_UsesVerb()
    {
        var msg = ErrorMessageBuilder.OperationFailed("render");
        Assert.Contains("render", msg, StringComparison.Ordinal);
        Assert.Contains("Failed to", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void OperationFailed_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.OperationFailed("export"));
    }

    // ─── ProcessingFailed ───────────────────────────────────────────────────────

    [Fact]
    public void ProcessingFailed_ReturnsExpectedSentinel()
    {
        var msg = ErrorMessageBuilder.ProcessingFailed();
        Assert.Equal(
            "An internal processing error occurred. Check inputs and retry.",
            msg);
    }

    [Fact]
    public void ProcessingFailed_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.ProcessingFailed());
    }

    // ─── UnsupportedFormat ──────────────────────────────────────────────────────

    [Fact]
    public void UnsupportedFormat_NoParam_ReturnsGenericSentinel()
    {
        var msg = ErrorMessageBuilder.UnsupportedFormat();
        Assert.Equal("The supplied file format is not supported.", msg);
    }

    [Fact]
    public void UnsupportedFormat_WithExtension_IncludesExtension()
    {
        var msg = ErrorMessageBuilder.UnsupportedFormat(".xls");
        Assert.Contains(".xls", msg, StringComparison.Ordinal);
        Assert.Contains("not supported", msg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UnsupportedFormat_NullExtension_ReturnsGeneric()
    {
        var msg = ErrorMessageBuilder.UnsupportedFormat();
        Assert.Equal("The supplied file format is not supported.", msg);
    }

    [Fact]
    public void UnsupportedFormat_HasNoLeakPatterns()
    {
        AssertNoLeakPatterns(ErrorMessageBuilder.UnsupportedFormat(".xlsx"));
    }

    // ─── ExtensionUnavailable ────────────────────────────────────────────────────

    [Fact]
    public void ExtensionUnavailable_ReturnsExpectedSentinel()
    {
        var msg = ErrorMessageBuilder.ExtensionUnavailable("test-tag");
        Assert.Equal("Extension is unavailable: test-tag.", msg);
    }

    [Fact]
    public void ExtensionUnavailable_IsNonEmpty()
    {
        Assert.NotEmpty(ErrorMessageBuilder.ExtensionUnavailable("test-tag"));
    }

    [Fact]
    public void ExtensionUnavailable_HasNoLeakPatterns()
    {
        var msg = ErrorMessageBuilder.ExtensionUnavailable("test-tag");
        Assert.DoesNotContain("Exception", msg, StringComparison.Ordinal);
        Assert.DoesNotContain("   at ", msg, StringComparison.Ordinal);
        Assert.DoesNotContain("StackTrace", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtensionUnavailable_IncludesTag()
    {
        var msg = ErrorMessageBuilder.ExtensionUnavailable("handshake-timeout");
        Assert.Contains("handshake-timeout", msg, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtensionUnavailable_DifferentTags_ProduceDifferentMessages()
    {
        var msg1 = ErrorMessageBuilder.ExtensionUnavailable("handshake-timeout");
        var msg2 = ErrorMessageBuilder.ExtensionUnavailable("initialization-failed");
        Assert.NotEqual(msg1, msg2);
    }

    // ─── cross-cutting: no path-separator leakage ────────────────────────────────

    [Fact]
    public void AllSentinels_DoNotContainForwardSlash()
    {
        // Sentinels must never carry file-system path separators.
        var sentinels = new[]
        {
            ErrorMessageBuilder.InvalidPassword(),
            ErrorMessageBuilder.IndexOutOfRange(),
            ErrorMessageBuilder.OutputDirectoryNotWritable(),
            ErrorMessageBuilder.OperationFailed(),
            ErrorMessageBuilder.ProcessingFailed(),
            ErrorMessageBuilder.UnsupportedFormat(),
            ErrorMessageBuilder.ExtensionUnavailable("test-tag")
        };

        foreach (var s in sentinels)
            Assert.DoesNotContain("/", s, StringComparison.Ordinal);
    }
}
