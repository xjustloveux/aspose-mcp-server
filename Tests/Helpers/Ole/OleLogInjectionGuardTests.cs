using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers.Ole;

namespace AsposeMcpServer.Tests.Helpers.Ole;

/// <summary>
///     AC-20 coverage: attacker-controlled strings carrying CRLF, NUL, ANSI CSI, and
///     C0/C1 controls must be stripped before any log-template interpolation or error
///     message emission. Validates <see cref="OleSanitizerHelper.SanitizeForLog" /> as
///     well as the builder-level guard in <see cref="OleErrorMessageBuilder" />.
/// </summary>
public class OleLogInjectionGuardTests
{
    /// <summary>
    ///     CRLF-laden raw names are stripped — no two-line log entry survives.
    /// </summary>
    [Fact]
    public void SanitizeForLog_StripsCrLf()
    {
        var raw = "line1\r\nFAKE LOG: admin login\r\nline3";
        var safe = OleSanitizerHelper.SanitizeForLog(raw);

        Assert.DoesNotContain('\r', safe);
        Assert.DoesNotContain('\n', safe);
        Assert.Contains("FAKE LOG", safe);
    }

    /// <summary>
    ///     ANSI CSI escape sequences (e.g. <c>\x1b[2J</c> clear-screen) never reach the
    ///     log sink.
    /// </summary>
    [Fact]
    public void SanitizeForLog_StripsAnsiEscapes()
    {
        var raw = "\u001b[2J\u001b[1;31mRED\u001b[0m text";
        var safe = OleSanitizerHelper.SanitizeForLog(raw);

        // The ESC (0x1B) control byte is the actual injection vector — a log consumer's
        // terminal interprets CSI sequences only when it sees the ESC prefix. The
        // sanitizer strips the ESC as a C0 control character; the residual literal
        // "[2J" text is harmless ASCII and is acceptable.
        Assert.DoesNotContain('\u001b', safe);
        Assert.Contains("RED", safe);
        Assert.Contains("text", safe);
    }

    /// <summary>
    ///     NUL and C0 controls are stripped.
    /// </summary>
    [Fact]
    public void SanitizeForLog_StripsNulAndC0Controls()
    {
        var raw = "a\0b\tc\u0001d";
        var safe = OleSanitizerHelper.SanitizeForLog(raw);

        Assert.DoesNotContain('\0', safe);
        Assert.DoesNotContain('\t', safe);
        Assert.DoesNotContain('\u0001', safe);
    }

    /// <summary>
    ///     Error-message builders route attacker-controlled content through
    ///     <see cref="OleSanitizerHelper.SanitizeForLog" /> — a raw filename containing
    ///     CRLF produces a single-line error message.
    /// </summary>
    [Fact]
    public void ErrorBuilder_SaveFailed_StripsCrLfFromFileName()
    {
        var attackerName = "legit.xlsx\r\nFAKE LOG LINE";
        var msg = OleErrorMessageBuilder.SaveFailed(attackerName);

        Assert.DoesNotContain('\r', msg);
        Assert.DoesNotContain('\n', msg);
    }

    /// <summary>
    ///     Unknown-operation errors sanitize the attacker-controlled operation token.
    /// </summary>
    [Fact]
    public void ErrorBuilder_UnknownOperation_StripsControlChars()
    {
        var op = "list\r\nadmin";
        var msg = OleErrorMessageBuilder.UnknownOperation(op);

        Assert.DoesNotContain('\r', msg);
        Assert.DoesNotContain('\n', msg);
    }

    /// <summary>
    ///     Invalid-path messages never leak the full path — only the sanitized basename
    ///     or a fixed sentinel — and are CRLF-clean.
    /// </summary>
    [Fact]
    public void ErrorBuilder_InvalidPath_DropsDirectoryComponents()
    {
        var attackerPath = "/etc/passwd\r\nfoo";
        var msg = OleErrorMessageBuilder.InvalidPath(attackerPath);

        Assert.DoesNotContain("/etc", msg);
        Assert.DoesNotContain('\r', msg);
        Assert.DoesNotContain('\n', msg);
    }
}
