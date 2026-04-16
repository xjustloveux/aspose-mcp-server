using Aspose.Cells;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

/// <summary>
///     Integration tests verifying that <see cref="UnprotectExcelHandler" /> routes
///     Aspose.Cells exceptions through <c>CellsErrorTranslator</c> and emits sanitized
///     BCL exceptions to callers. Regression guard for the unified error translator.
/// </summary>
public class UnprotectExcelHandlerTranslatorTests : ExcelHandlerTestBase
{
    private readonly UnprotectExcelHandler _handler = new();

    /// <summary>
    ///     When the supplied password for a protected worksheet is wrong, Aspose throws a
    ///     <see cref="Aspose.Cells.CellsException" /> with
    ///     <see cref="ExceptionType.IncorrectPassword" />. After translation the handler must
    ///     re-throw an <see cref="UnauthorizedAccessException" /> with the fixed
    ///     InvalidPassword sentinel — NOT the raw Aspose message.
    /// </summary>
    [Fact]
    public void Execute_WrongWorksheetPassword_ThrowsUnauthorizedAccessException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "correct", null);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "password", "wrong" }
        });

        var ex = Assert.Throws<UnauthorizedAccessException>(() => _handler.Execute(context, parameters));

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), ex.Message);
    }

    /// <summary>
    ///     The translated exception must carry the fixed sentinel message and must NOT
    ///     contain any text from the original Aspose exception (path leakage guard).
    /// </summary>
    [Fact]
    public void Execute_WrongWorksheetPassword_MessageIsFixedSentinel_NoRawLeakage()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "secret_password", null);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "password", "wrong_attempt" }
        });

        var ex = Assert.Throws<UnauthorizedAccessException>(() => _handler.Execute(context, parameters));

        // The supplied passwords must not appear in the translated message.
        Assert.DoesNotContain("secret_password", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("wrong_attempt", ex.Message, StringComparison.Ordinal);
        // No inner exception must be attached (Aspose frames must not escape).
        Assert.Null(ex.InnerException);
    }
}
