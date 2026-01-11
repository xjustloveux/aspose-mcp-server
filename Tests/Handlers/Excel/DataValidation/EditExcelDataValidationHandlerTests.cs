using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataValidation;

public class EditExcelDataValidationHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelDataValidationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_EditsOnCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Opt1,Opt2");
        AddValidationToSheet(workbook.Worksheets[1], "B1", "X,Y,Z");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "validationIndex", 0 },
            { "formula1", "A,B,C" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Opt1,Opt2", workbook.Worksheets[0].Validations[0].Formula1);
        Assert.Equal("A,B,C", workbook.Worksheets[1].Validations[0].Formula1);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsValidationType()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "validationType", "WholeNumber" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Edited", result);
        Assert.Equal(ValidationType.WholeNumber, workbook.Worksheets[0].Validations[0].Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsFormula1()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "formula1", "NewA,NewB,NewC" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formula1=NewA,NewB,NewC", result);
        Assert.Equal("NewA,NewB,NewC", workbook.Worksheets[0].Validations[0].Formula1);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsFormula2()
    {
        var workbook = CreateWorkbookWithWholeNumberValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "formula2", "200" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formula2=200", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsInCellDropDown()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "inCellDropDown", false }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[0].Validations[0].InCellDropDown);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsErrorMessage()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "errorMessage", "New error message" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("New error message", workbook.Worksheets[0].Validations[0].ErrorMessage);
        Assert.True(workbook.Worksheets[0].Validations[0].ShowError);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsInputMessage()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "inputMessage", "New input message" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("New input message", workbook.Worksheets[0].Validations[0].InputMessage);
        Assert.True(workbook.Worksheets[0].Validations[0].ShowInput);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultipleChanges_AppliesAll()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "formula1", "X,Y,Z" },
            { "errorMessage", "Error!" },
            { "inCellDropDown", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formula1", result);
        Assert.Contains("ErrorMessage", result);
        Assert.Contains("InCellDropDown", result);
        Assert.Equal("X,Y,Z", workbook.Worksheets[0].Validations[0].Formula1);
        Assert.Equal("Error!", workbook.Worksheets[0].Validations[0].ErrorMessage);
        Assert.True(workbook.Worksheets[0].Validations[0].InCellDropDown);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutValidationIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithValidation();
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
        var workbook = CreateWorkbookWithValidation();
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
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "validationIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChanges()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No changes", result);
    }

    [Fact]
    public void Execute_ReturnsValidationIndex()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "formula1", "A,B" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithValidation()
    {
        var workbook = new Workbook();
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Opt1,Opt2,Opt3");
        return workbook;
    }

    private static Workbook CreateWorkbookWithWholeNumberValidation()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 };
        var index = sheet.Validations.Add(area);
        var validation = sheet.Validations[index];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";
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
