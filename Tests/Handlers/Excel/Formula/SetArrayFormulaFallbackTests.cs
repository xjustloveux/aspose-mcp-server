using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

/// <summary>
///     Covers TEST-08. Setting an array formula tries three increasingly permissive strategies and
///     only the first one had tests, so the two fallbacks were untested code that runs exactly when
///     something has already gone wrong. Each layer is driven directly here, and the chain is
///     checked to fall through in order.
/// </summary>
public class SetArrayFormulaFallbackTests : ExcelHandlerTestBase
{
    /// <summary>Builds the context the three strategies operate on.</summary>
    /// <param name="worksheet">Worksheet holding the range.</param>
    /// <param name="range">Range in A1 notation.</param>
    /// <param name="formula">Formula text, with or without a leading equals sign.</param>
    /// <returns>The populated context.</returns>
    private static SetArrayFormulaHandler.FormulaContext CreateContext(
        Worksheet worksheet, string range, string formula)
    {
        var rangeObject = worksheet.Cells.CreateRange(range);
        return new SetArrayFormulaHandler.FormulaContext(worksheet, rangeObject, formula, range);
    }

    [Fact]
    public void PrimaryMethod_ShouldSetATrueArrayFormula()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);

        var result = SetArrayFormulaHandler.TryPrimaryMethod(CreateContext(sheet, "C1:C2", "=A1:A2*2"));

        Assert.True(result.IsSuccess, result.Message);
        Assert.True(sheet.Cells["C1"].IsArrayFormula);
    }

    [Fact]
    public void AlternativeMethod_ShouldReportWhetherItProducedAnArrayFormula()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);

        var result = SetArrayFormulaHandler.TryAlternativeMethod(CreateContext(sheet, "D1:D2", "A1:A2*2"));

        // The five-parameter overload either produces an array formula, reports that it did
        // not, or throws internally; all three are defined outcomes and none may escape.
        // Measured on Aspose.Cells 23.10 for this input: it throws, so the layer reports
        // "alternative method failed" and the chain moves on.
        Assert.Equal(sheet.Cells["D1"].IsArrayFormula, result.IsSuccess);
        if (!result.IsSuccess)
            Assert.Contains(result.Message, new[]
            {
                "alternative method did not set array formula",
                "alternative method failed"
            });
    }

    [Fact]
    public void FallbackMethod_ShouldWriteAPlainFormulaToEveryCell()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);

        var result = SetArrayFormulaHandler.TryFallbackMethod(CreateContext(sheet, "E1:E2", "SUM(A1:A2)"));

        Assert.True(result.IsSuccess, result.Message);
        Assert.Contains("not a true array formula", result.Message);
        Assert.Equal("=SUM(A1:A2)", sheet.Cells["E1"].Formula);
        Assert.Equal("=SUM(A1:A2)", sheet.Cells["E2"].Formula);
        Assert.False(sheet.Cells["E1"].IsArrayFormula);
    }

    [Fact]
    public void FallbackMethod_ShouldAddTheEqualsSignOnlyOnce()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];

        var result = SetArrayFormulaHandler.TryFallbackMethod(CreateContext(sheet, "F1:F1", "=1+1"));

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal("=1+1", sheet.Cells["F1"].Formula);
    }

    [Fact]
    public void FallbackMethod_ShouldReportFailureInsteadOfThrowing()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];

        var result = SetArrayFormulaHandler.TryFallbackMethod(CreateContext(sheet, "G1:G2", "=THIS IS NOT A FORMULA("));

        Assert.False(result.IsSuccess);
        Assert.Equal("fallback method failed", result.Message);
    }

    [Fact]
    public void Chain_ShouldReportEveryFailureWhenNoStrategyWorks()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];

        var result = SetArrayFormulaHandler.TrySetArrayFormula(
            CreateContext(sheet, "H1:H2", "=THIS IS NOT A FORMULA("));

        Assert.False(result.IsSuccess);
        Assert.Contains("primary method failed", result.Message);
        Assert.Contains("fallback method failed", result.Message);
    }

    [Fact]
    public void Chain_ShouldStopAtTheFirstStrategyThatWorks()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);

        var result = SetArrayFormulaHandler.TrySetArrayFormula(CreateContext(sheet, "I1:I2", "=A1:A2*2"));

        Assert.True(result.IsSuccess, result.Message);
        Assert.DoesNotContain("not a true array formula", result.Message);
        Assert.True(sheet.Cells["I1"].IsArrayFormula);
    }
}
