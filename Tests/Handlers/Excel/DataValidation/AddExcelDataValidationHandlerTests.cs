using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataValidation;

public class AddExcelDataValidationHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelDataValidationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AddsToCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1" },
            { "validationType", "List" },
            { "formula1", "X,Y,Z" }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].Validations);
        Assert.Single(workbook.Worksheets[1].Validations);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsListValidation()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A10" },
            { "validationType", "List" },
            { "formula1", "Option1,Option2,Option3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Data validation added", result);
        Assert.Single(workbook.Worksheets[0].Validations);
        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal(ValidationType.List, validation.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsWholeNumberValidation()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B10" },
            { "validationType", "WholeNumber" },
            { "formula1", "1" },
            { "formula2", "100" },
            { "operatorType", "Between" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Data validation added", result);
        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Contains("1", validation.Formula1);
        Assert.Contains("100", validation.Formula2);
        AssertModified(context);
    }

    [Theory]
    [InlineData("List", ValidationType.List)]
    [InlineData("WholeNumber", ValidationType.WholeNumber)]
    [InlineData("Decimal", ValidationType.Decimal)]
    [InlineData("Date", ValidationType.Date)]
    [InlineData("TextLength", ValidationType.TextLength)]
    public void Execute_WithVariousValidationTypes_SetsCorrectType(string typeStr, ValidationType expectedType)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", typeStr },
            { "formula1", "1" }
        });

        _handler.Execute(context, parameters);

        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal(expectedType, validation.Type);
    }

    #endregion

    #region Optional Parameters

    [Fact]
    public void Execute_WithErrorMessage_SetsErrorMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", "List" },
            { "formula1", "Yes,No" },
            { "errorMessage", "Please select Yes or No" }
        });

        _handler.Execute(context, parameters);

        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal("Please select Yes or No", validation.ErrorMessage);
        Assert.True(validation.ShowError);
    }

    [Fact]
    public void Execute_WithInputMessage_SetsInputMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", "List" },
            { "formula1", "Yes,No" },
            { "inputMessage", "Select an option" }
        });

        _handler.Execute(context, parameters);

        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal("Select an option", validation.InputMessage);
        Assert.True(validation.ShowInput);
    }

    [Fact]
    public void Execute_WithInCellDropDownFalse_DisablesDropdown()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", "List" },
            { "formula1", "A,B,C" },
            { "inCellDropDown", false }
        });

        _handler.Execute(context, parameters);

        var validation = workbook.Worksheets[0].Validations[0];
        Assert.False(validation.InCellDropDown);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationType", "List" },
            { "formula1", "A,B" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutValidationType_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "formula1", "A,B" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("validationType", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutFormula1_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", "List" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("formula1", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "range", "A1" },
            { "validationType", "List" },
            { "formula1", "A,B" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRangeInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C5:C20" },
            { "validationType", "List" },
            { "formula1", "Yes,No" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("C5:C20", result);
    }

    [Fact]
    public void Execute_ReturnsValidationTypeInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "validationType", "WholeNumber" },
            { "formula1", "1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("WholeNumber", result);
    }

    #endregion
}
