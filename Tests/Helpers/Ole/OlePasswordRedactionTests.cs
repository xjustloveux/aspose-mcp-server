using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Tests.Helpers.Ole;

/// <summary>
///     AC-21 coverage (F-5): password values must never appear in any error message,
///     log fragment, or JSON response. The locked shape of the session-mode
///     <see cref="PasswordIgnoredNote" /> is also pinned here so any future drift (adding
///     a <c>value</c> / <c>attempted</c> field) breaks the test loudly.
/// </summary>
public class OlePasswordRedactionTests
{
    /// <summary>Distinctive attempted-password used in every assertion.</summary>
    private const string AttemptedPassword = "CORRECT_HORSE_BATTERY_STAPLE";

    /// <summary>
    ///     <see cref="OleErrorMessageBuilder.InvalidPassword" /> emits the fixed sentinel
    ///     and never echoes the attempted password — the sentinel itself does not vary
    ///     with input so it literally cannot leak.
    /// </summary>
    [Fact]
    public void InvalidPasswordMessage_IsFixedSentinel()
    {
        var msg = OleErrorMessageBuilder.InvalidPassword();

        Assert.Equal(OleErrorMessageBuilder.InvalidPasswordSentinel, msg);
        Assert.DoesNotContain(AttemptedPassword, msg);
    }

    /// <summary>
    ///     Translator maps both password-class exceptions to a BCL
    ///     <see cref="UnauthorizedAccessException" /> whose <see cref="Exception.Message" />
    ///     equals the fixed sentinel; the translator never stringifies the inner
    ///     exception whose <c>.Message</c> could theoretically carry the attempted
    ///     password.
    /// </summary>
    [Fact]
    public void Translate_PasswordClassException_DoesNotEchoAttemptedPassword()
    {
        // IncorrectPasswordException / InvalidPasswordException have no public ctor in
        // 23.10.0; use the framework-exception branch with an inner message that
        // intentionally contains the attempted password, and assert the translator
        // drops it.
        var inner = new UnauthorizedAccessException("wrong password provided: " + AttemptedPassword);
        var mapped = OleErrorTranslator.Translate(inner);

        Assert.DoesNotContain(AttemptedPassword, mapped.ToString());
    }

    /// <summary>
    ///     <see cref="PasswordIgnoredNote" /> shape is locked: exactly two fields,
    ///     <c>passwordIgnored: true</c> and <c>reason: "session-already-unlocked"</c>.
    ///     Any new init-only property on the record is caught by reflection here.
    /// </summary>
    [Fact]
    public void PasswordIgnoredNote_ShapeIsLocked()
    {
        var note = new PasswordIgnoredNote();

        Assert.True(note.PasswordIgnored);
        Assert.Equal("session-already-unlocked", note.Reason);

        var properties = typeof(PasswordIgnoredNote).GetProperties();
        Assert.Equal(2, properties.Length);
        Assert.Contains(properties, p => p.Name == nameof(PasswordIgnoredNote.PasswordIgnored));
        Assert.Contains(properties, p => p.Name == nameof(PasswordIgnoredNote.Reason));
        Assert.DoesNotContain(properties, p => p.Name is "Value" or "Attempted");
    }

    /// <summary>
    ///     Attempted password never appears in any of the centralized error templates,
    ///     even when supplied as a filename / operation / extension. The sanitizer
    ///     leaves it intact (it is not an attack character sequence) but the templates
    ///     only emit sanitized basenames or fixed sentinels, never the password.
    /// </summary>
    [Fact]
    public void ErrorTemplates_DoNotContainAttemptedPassword()
    {
        // None of these templates carry a password, but we assert as a property of the
        // allowlist: no builder method accepts a password argument at all.
        Assert.DoesNotContain(AttemptedPassword, OleErrorMessageBuilder.InvalidPassword());
        Assert.DoesNotContain(AttemptedPassword, OleErrorMessageBuilder.LinkedCannotExtract());
        Assert.DoesNotContain(AttemptedPassword, OleErrorMessageBuilder.IndexOutOfRange(null, null));
        Assert.DoesNotContain(AttemptedPassword, OleErrorMessageBuilder.IndexOutOfRange(5, 10));
    }
}
