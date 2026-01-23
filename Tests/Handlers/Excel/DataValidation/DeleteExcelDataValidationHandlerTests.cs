using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataValidation;

public class DeleteExcelDataValidationHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelDataValidationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Yes,No");
        AddValidationToSheet(workbook.Worksheets[1], "B1", "X,Y,Z");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "validationIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].Validations);
        Assert.Empty(workbook.Worksheets[1].Validations);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var workbook = CreateWorkbookWithValidations(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("remaining: 2", result.Message);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesValidation()
    {
        var workbook = CreateWorkbookWithValidations(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Single(workbook.Worksheets[0].Validations);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesValidationAtVariousIndices(int index)
    {
        var workbook = CreateWorkbookWithValidations(3);
        var initialCount = workbook.Worksheets[0].Validations.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", index }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, workbook.Worksheets[0].Validations.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutValidationIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithValidations(1);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("validationIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidValidationIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateWorkbookWithValidations(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithValidations(1);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "validationIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithValidations(int count)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < count; i++) AddValidationToSheet(sheet, $"A{i + 1}", $"Option{i}A,Option{i}B");
        return workbook;
    }

    private static void AddValidationToSheet(Worksheet sheet, string cellAddress, string listValues)
    {
        var cell = sheet.Cells[cellAddress];
        var area = new CellArea
        {
            StartRow = cell.Row,
            StartColumn = cell.Column,
            EndRow = cell.Row,
            EndColumn = cell.Column
        };
        var index = sheet.Validations.Add(area);
        var validation = sheet.Validations[index];
        validation.Type = ValidationType.List;
        validation.Formula1 = listValues;
    }

    #endregion
}
