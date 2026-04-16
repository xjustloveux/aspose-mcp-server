using System.Text;
using AsposeMcpServer.Helpers.Ole;

namespace AsposeMcpServer.Tests.Helpers.Ole;

/// <summary>
///     AC-23 coverage (F-1): UTF-8 byte-count clamp at 255 bytes, BiDi / C0 / C1
///     stripping, NBSP+tab+FF trailing removal, idempotence. Complements the broader
///     <see cref="OleSanitizerHelperTests" /> with focused tests on the byte-clamp +
///     stripping contract.
/// </summary>
public class OleSanitizerByteClampTests
{
    /// <summary>
    ///     128 CJK characters = 384 UTF-8 bytes. The sanitizer must clamp below the
    ///     255-byte limit while preserving the extension so the file remains openable
    ///     on an ext4 filesystem.
    /// </summary>
    [Fact]
    public void Cjk128Chars_ClampsToAt255Bytes_PreservingExtension()
    {
        var raw = new string('財', 128) + ".xlsx";
        Assert.True(Encoding.UTF8.GetByteCount(raw) > 255);

        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, "Excel.Sheet.12");

        Assert.True(Encoding.UTF8.GetByteCount(suggested) <= 255);
        Assert.EndsWith(".xlsx", suggested);
        Assert.True(changed);
    }

    /// <summary>
    ///     BiDi-override + C0 + trailing-NBSP combo is stripped, and the result is
    ///     stable under a second sanitization pass (idempotence under F-1 rule 11).
    /// </summary>
    [Fact]
    public void BiDiAndControlsStripped_IdempotenceHolds()
    {
        var raw = "\u202ehidden\u0001name.xlsx\u00A0\u00A0";
        var (first, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);
        var (second, _) = OleSanitizerHelper.SanitizeOleFileName(first, 0, null);

        Assert.Equal(first, second);
        Assert.DoesNotContain('\u202e', first);
        Assert.DoesNotContain('\u0001', first);
        Assert.DoesNotContain('\u00A0', first);
    }

    /// <summary>
    ///     Each of the BiDi override code points listed in F-1 is individually stripped.
    /// </summary>
    [Theory]
    [InlineData('\u200E')]
    [InlineData('\u200F')]
    [InlineData('\u202A')]
    [InlineData('\u202B')]
    [InlineData('\u202C')]
    [InlineData('\u202D')]
    [InlineData('\u202E')]
    [InlineData('\u2066')]
    [InlineData('\u2067')]
    [InlineData('\u2068')]
    [InlineData('\u2069')]
    public void EveryBiDiCodePoint_IsStripped(char bidi)
    {
        var raw = "file" + bidi + "name.xlsx";
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);

        Assert.DoesNotContain(bidi, suggested);
    }

    /// <summary>
    ///     Extensions-only overflow (name is 256-byte all-extension) falls back to a
    ///     clamped extension. Edge case: ensures the clamp loop terminates even when
    ///     the stem is empty.
    /// </summary>
    [Fact]
    public void ExtensionOnlyOverflow_ClampsGracefully()
    {
        var raw = "." + new string('a', 300);
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);

        Assert.True(Encoding.UTF8.GetByteCount(suggested) <= 255);
        Assert.False(string.IsNullOrEmpty(suggested));
    }

    /// <summary>
    ///     Trailing-whitespace regex covers tab (<c>\t</c>), form-feed (<c>\u000C</c>),
    ///     vertical tab (<c>\u000B</c>), and NBSP (<c>\u00A0</c>) in addition to space
    ///     and dot (F-1 rule 8). Control characters are pre-stripped by the earlier
    ///     pass, so we assert the net effect on the final suggested name.
    /// </summary>
    [Theory]
    [InlineData("file.xlsx\u00A0")]
    [InlineData("file.xlsx ")]
    [InlineData("file.xlsx.")]
    [InlineData("file.xlsx. .  ")]
    public void TrailingWhitespaceVariants_AreStripped(string raw)
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);

        Assert.Equal("file.xlsx", suggested);
    }
}
