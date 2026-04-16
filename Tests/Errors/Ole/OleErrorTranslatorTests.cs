using AsposeMcpServer.Errors.Ole;

namespace AsposeMcpServer.Tests.Errors.Ole;

/// <summary>
///     Unit tests for <see cref="OleErrorTranslator" />. Verifies that Aspose-specific and
///     framework exceptions map to the expected sanitized BCL exception types and that
///     no raw inner-exception text leaks through (F-10).
/// </summary>
public class OleErrorTranslatorTests
{
    // Note: Aspose.Words.IncorrectPasswordException and Aspose.Slides.InvalidPasswordException
    // do not expose public constructors in 23.10.0, so the password-class mapping is exercised
    // indirectly via integration tests (test-engineer stage). The translator's fallback and
    // framework-exception branches are exercised here.

    [Fact]
    public void Translate_IoException_MapsToIOException()
    {
        var ex = new IOException("disk full");
        var result = OleErrorTranslator.Translate(ex, "report.xlsx");

        Assert.IsType<IOException>(result);
        Assert.Contains("report.xlsx", result.Message);
        Assert.DoesNotContain("disk full", result.Message);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_MapsToUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = OleErrorTranslator.Translate(ex, "target");

        Assert.IsType<UnauthorizedAccessException>(result);
        Assert.DoesNotContain("permission denied", result.Message);
        Assert.DoesNotContain("/secret/path", result.Message);
    }

    [Fact]
    public void Translate_ArgumentOutOfRange_MapsToArgumentOutOfRangeException()
    {
        // ReSharper disable once NotResolvedInText — "index" is the param name the simulated Aspose caller would supply
        var ex = new ArgumentOutOfRangeException("index", 42, "raw");
        var result = OleErrorTranslator.Translate(ex);

        Assert.IsType<ArgumentOutOfRangeException>(result);
        Assert.DoesNotContain("raw", result.Message);
    }

    [Fact]
    public void Translate_GenericException_MapsToInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = OleErrorTranslator.Translate(ex, "x");

        Assert.IsType<InvalidOperationException>(result);
        Assert.DoesNotContain("boom", result.Message);
    }
}
