using System.Reflection;
using Aspose.Cells;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Errors.Excel;

namespace AsposeMcpServer.Tests.Errors.Excel;

/// <summary>
///     Unit tests for <see cref="CellsErrorTranslator" />. Verifies that Aspose.Cells and BCL
///     exceptions map to the expected sanitized BCL exception types and that no raw
///     inner-exception text (including file paths) leaks through (charter §5 red-line F-10).
/// </summary>
/// <remarks>
///     <see cref="CellsException" /> has no public constructors in Aspose.Cells 23.10.0, so
///     instances are created via private-constructor reflection — analogous to the comment in
///     <c>OleErrorTranslatorTests</c> for Aspose.Words/Slides password exceptions.
/// </remarks>
public class CellsErrorTranslatorTests
{
    // ─── factory ──────────────────────────────────────────────────────────────────

    /// <summary>
    ///     Creates a <see cref="CellsException" /> via private-constructor reflection.
    ///     Aspose.Cells 23.10.0 exposes no public constructor; the private signature is
    ///     <c>(ExceptionType, string)</c>.
    /// </summary>
    private static CellsException MakeCellsException(ExceptionType code, string message)
    {
        var ctor = typeof(CellsException)
                       .GetConstructor(
                           BindingFlags.NonPublic | BindingFlags.Instance,
                           null,
                           [typeof(ExceptionType), typeof(string)],
                           null)
                   ?? throw new InvalidOperationException(
                       "CellsException private ctor (ExceptionType, string) not found — Aspose version changed?");

        return (CellsException)ctor.Invoke([code, message]);
    }

    // ─── CellsException mappings ─────────────────────────────────────────────────

    [Fact]
    public void Translate_IncorrectPassword_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeCellsException(ExceptionType.IncorrectPassword, "wrong pw");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_IncorrectPassword_MessageIsInvalidPasswordSentinel()
    {
        var ex = MakeCellsException(ExceptionType.IncorrectPassword, "secret internal detail");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), result.Message);
    }

    [Fact]
    public void Translate_IncorrectPassword_DoesNotLeakInnerMessage()
    {
        var ex = MakeCellsException(ExceptionType.IncorrectPassword, "secret internal detail");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.DoesNotContain("secret internal detail", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_InvalidData_ReturnsInvalidOperationException()
    {
        var ex = MakeCellsException(ExceptionType.InvalidData, "corrupted");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_InvalidData_MessageIsProcessingFailedSentinel()
    {
        var ex = MakeCellsException(ExceptionType.InvalidData, "corrupted");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    [Fact]
    public void Translate_FileCorrupted_ReturnsInvalidOperationException()
    {
        var ex = MakeCellsException(ExceptionType.FileCorrupted, "bad bytes");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_FileFormat_ReturnsNotSupportedException()
    {
        var ex = MakeCellsException(ExceptionType.FileFormat, "wrong format");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<NotSupportedException>(result);
    }

    [Fact]
    public void Translate_UnsupportedFeature_ReturnsNotSupportedException()
    {
        var ex = MakeCellsException(ExceptionType.UnsupportedFeature, "not supported");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<NotSupportedException>(result);
    }

    [Fact]
    public void Translate_Limitation_ReturnsNotSupportedException()
    {
        var ex = MakeCellsException(ExceptionType.Limitation, "limit reached");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<NotSupportedException>(result);
    }

    [Fact]
    public void Translate_IoCode_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeCellsException(ExceptionType.IO, "disk error");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_Permission_ReturnsUnauthorizedAccessException()
    {
        var ex = MakeCellsException(ExceptionType.Permission, "perm denied");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_IoCode_WithContextBasename_IncludesContextInMessage()
    {
        var ex = MakeCellsException(ExceptionType.IO, "disk error");
        var result = CellsErrorTranslator.Translate(ex, "report.xlsx");

        Assert.Contains("report.xlsx", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_IoCode_DoesNotLeakInnerMessage()
    {
        var ex = MakeCellsException(ExceptionType.IO, "disk error at /secret/path");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.DoesNotContain("disk error", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
    }

    // ─── BCL exception mappings ──────────────────────────────────────────────────

    [Fact]
    public void Translate_ArgumentOutOfRange_ReturnsArgumentOutOfRangeException()
    {
        // ReSharper disable once NotResolvedInText — "index" is the param name the simulated Aspose caller would supply
        var ex = new ArgumentOutOfRangeException("index", 42, "raw detail");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<ArgumentOutOfRangeException>(result);
    }

    [Fact]
    public void Translate_ArgumentOutOfRange_MessageIsIndexOutOfRangeSentinel()
    {
        // ReSharper disable once NotResolvedInText — "index" is the param name the simulated Aspose caller would supply
        var ex = new ArgumentOutOfRangeException("index", 42, "raw detail");
        var result = Assert.IsType<ArgumentOutOfRangeException>(CellsErrorTranslator.Translate(ex));

        // .NET appends "(Parameter 'ex')" to Message when ParamName is set; sentinel must be the prefix.
        var expectedSentinel = ErrorMessageBuilder.IndexOutOfRange();
        Assert.StartsWith(expectedSentinel, result.Message, StringComparison.Ordinal);
        Assert.Equal("ex", result.ParamName);
    }

    [Fact]
    public void Translate_ArgumentOutOfRange_DoesNotLeakRawDetail()
    {
        // ReSharper disable once NotResolvedInText — "index" is the param name the simulated Aspose caller would supply
        var ex = new ArgumentOutOfRangeException("index", 42, "raw detail");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.DoesNotContain("raw detail", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_ReturnsUnauthorizedAccessException()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_MessageIsOutputDirectoryNotWritableSentinel()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.OutputDirectoryNotWritable(), result.Message);
    }

    [Fact]
    public void Translate_UnauthorizedAccess_DoesNotLeakPath()
    {
        var ex = new UnauthorizedAccessException("permission denied /secret/path");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/secret/path", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("permission denied", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_DirectoryNotFound_ReturnsUnauthorizedAccessException()
    {
        var ex = new DirectoryNotFoundException("no dir at /hidden/output");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<UnauthorizedAccessException>(result);
    }

    [Fact]
    public void Translate_GenericException_ReturnsInvalidOperationException()
    {
        var ex = new InvalidOperationException("boom");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.IsType<InvalidOperationException>(result);
    }

    [Fact]
    public void Translate_GenericException_MessageIsProcessingFailedSentinel()
    {
        var ex = new InvalidOperationException("boom");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), result.Message);
    }

    // ─── Security: file-path leakage ────────────────────────────────────────────

    [Fact]
    public void Translate_GenericExceptionWithFilePath_DoesNotLeakPath()
    {
        var ex = new Exception("secret path /etc/passwd could not be read");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.DoesNotContain("/etc/passwd", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("secret path", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Translate_GenericExceptionWithWindowsPath_DoesNotLeakPath()
    {
        var ex = new Exception(@"C:\Users\admin\Documents\private.xlsx not found");
        var result = CellsErrorTranslator.Translate(ex);

        // None of the internal path components should appear in the output.
        Assert.DoesNotContain("admin", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private.xlsx", result.Message, StringComparison.Ordinal);
    }

    // ─── No inner-exception attached ─────────────────────────────────────────────

    [Fact]
    public void Translate_NeverAttachesInnerException()
    {
        // The returned exception must never carry the original as InnerException, to
        // prevent Aspose stack frames from escaping via serialization.
        var ex = MakeCellsException(ExceptionType.IncorrectPassword, "internal detail");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }

    [Fact]
    public void Translate_Generic_NeverAttachesInnerException()
    {
        var ex = new Exception("boom");
        var result = CellsErrorTranslator.Translate(ex);

        Assert.Null(result.InnerException);
    }
}
