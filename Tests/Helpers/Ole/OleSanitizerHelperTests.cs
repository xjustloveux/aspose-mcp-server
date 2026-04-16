using System.Text;
using AsposeMcpServer.Helpers.Ole;

namespace AsposeMcpServer.Tests.Helpers.Ole;

/// <summary>
///     Unit tests for <see cref="OleSanitizerHelper" />. Covers the F-1 amendments:
///     BiDi / C0 / C1 stripping, trailing NBSP/tab/FF trimming, UTF-8 byte clamp,
///     idempotence, and the core traversal / NUL / reserved-name rules.
/// </summary>
public class OleSanitizerHelperTests
{
    [Fact]
    public void SanitizeOleFileName_NullInput_ReturnsFallbackUsingProgId()
    {
        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName(null, 7, "Excel.Sheet.12");

        Assert.Equal("ole_7.xlsx", suggested);
        Assert.True(changed);
    }

    [Fact]
    public void SanitizeOleFileName_EmptyProgId_FallsBackToBin()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(null, 0, null);

        Assert.Equal("ole_0.bin", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_SimpleName_PassesThroughUnchanged()
    {
        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName("report.xlsx", 0, "Excel.Sheet.12");

        Assert.Equal("report.xlsx", suggested);
        Assert.False(changed);
    }

    [Fact]
    public void SanitizeOleFileName_PathTraversal_StripsSeparatorsAndDotDot()
    {
        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName("..\\..\\etc\\passwd", 0, null);

        Assert.DoesNotContain("..", suggested);
        Assert.DoesNotContain("\\", suggested);
        Assert.DoesNotContain("/", suggested);
        Assert.True(changed);
    }

    [Fact]
    public void SanitizeOleFileName_UncPrefix_StripsHostComponent()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("\\\\attacker\\share\\x.exe", 0, null);

        Assert.Equal("x.exe", suggested);
        Assert.DoesNotContain("\\", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_AbsoluteWindowsPath_StripsDriveAndFolders()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("C:\\Users\\victim\\x.xlsx", 0, null);

        Assert.Equal("x.xlsx", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_NullByte_IsReplaced()
    {
        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName("file\0name.xlsx", 0, null);

        Assert.DoesNotContain('\0', suggested);
        Assert.True(changed);
    }

    [Fact]
    public void SanitizeOleFileName_ReservedNameCon_IsPrefixed()
    {
        var (suggested, changed) = OleSanitizerHelper.SanitizeOleFileName("CON.xlsx", 0, null);

        Assert.True(OleSanitizerHelper.IsWindowsReservedName("CON.xlsx"));
        Assert.StartsWith("_", suggested);
        Assert.True(changed);
    }

    [Fact]
    public void SanitizeOleFileName_RtloOverride_IsStripped()
    {
        // "photo\u202egpj.exe" would render right-to-left as "photoexe.jpg"
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("photo\u202egpj.exe", 0, null);

        Assert.DoesNotContain('\u202e', suggested);
    }

    [Fact]
    public void SanitizeOleFileName_NbspTrailing_IsStripped()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("file.xlsx\u00A0\u00A0", 0, null);

        Assert.Equal("file.xlsx", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_TrailingDotsAndSpaces_AreStripped()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("file.xlsx. .", 0, null);

        Assert.Equal("file.xlsx", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_CjkExceeds255Bytes_ClampsByUtf8ByteCount()
    {
        // 128 CJK chars = 384 bytes; after ".xlsx" extension must still be ≤ 255 bytes.
        var stem = new string('財', 128);
        var raw = stem + ".xlsx";

        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);

        Assert.True(Encoding.UTF8.GetByteCount(suggested) <= 255,
            $"Expected ≤255 UTF-8 bytes, got {Encoding.UTF8.GetByteCount(suggested)}");
        Assert.EndsWith(".xlsx", suggested);
    }

    [Fact]
    public void SanitizeOleFileName_IsIdempotent()
    {
        var raw = "\u202e..\\..\\CON.xlsx\u00A0";
        var (first, _) = OleSanitizerHelper.SanitizeOleFileName(raw, 0, null);
        var (second, _) = OleSanitizerHelper.SanitizeOleFileName(first, 0, null);

        Assert.Equal(first, second);
    }

    [Fact]
    public void SanitizeOleFileName_ControlCharactersAreStripped()
    {
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName("a\rb\nc\td.xlsx", 0, null);

        Assert.DoesNotContain('\r', suggested);
        Assert.DoesNotContain('\n', suggested);
        Assert.DoesNotContain('\t', suggested);
    }

    [Theory]
    [InlineData("Excel.Sheet.12", ".xlsx")]
    [InlineData("excel.sheet.12", ".xlsx")]
    [InlineData("Excel.Sheet.8", ".xls")]
    [InlineData("Word.Document.12", ".docx")]
    [InlineData("Word.Document.8", ".doc")]
    [InlineData("PowerPoint.Show.12", ".pptx")]
    [InlineData("AcroExch.Document", ".pdf")]
    [InlineData("Package", ".bin")]
    [InlineData(null, ".bin")]
    [InlineData("", ".bin")]
    [InlineData("UnknownProgId.42", ".bin")]
    public void ExtensionFromProgId_MapsExpected(string? progId, string expected)
    {
        Assert.Equal(expected, OleSanitizerHelper.ExtensionFromProgId(progId));
    }

    [Theory]
    [InlineData("xlsx", ".xlsx")]
    [InlineData(".xlsx", ".xlsx")]
    [InlineData("", "")]
    [InlineData(null, "")]
    [InlineData("   ", "")]
    public void NormalizeExtension_MapsExpected(string? input, string expected)
    {
        Assert.Equal(expected, OleSanitizerHelper.NormalizeExtension(input));
    }

    [Fact]
    public void SanitizeForLog_StripsCrLfAndAnsi()
    {
        var raw = "\r\nFAKE LOG: admin login\r\n\u001b[2Jsome text";
        var safe = OleSanitizerHelper.SanitizeForLog(raw);

        Assert.DoesNotContain('\r', safe);
        Assert.DoesNotContain('\n', safe);
        Assert.DoesNotContain('\u001b', safe);
        Assert.Contains("FAKE LOG", safe);
        Assert.Contains("some text", safe);
    }

    [Fact]
    public void SanitizeForLog_NullReturnsEmpty()
    {
        Assert.Equal(string.Empty, OleSanitizerHelper.SanitizeForLog(null));
    }

    [Theory]
    [InlineData("CON", true)]
    [InlineData("con.xlsx", true)]
    [InlineData("COM1.txt", true)]
    [InlineData("lpt9", true)]
    [InlineData("console", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void IsWindowsReservedName_MatchesExpected(string? input, bool expected)
    {
        Assert.Equal(expected, OleSanitizerHelper.IsWindowsReservedName(input));
    }
}
