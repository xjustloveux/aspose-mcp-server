using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataValidation;

public class SetMessagesExcelDataValidationHandlerTests : ExcelHandlerTestBase
{
    private readonly SetMessagesExcelDataValidationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetMessages()
    {
        Assert.Equal("set_messages", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_UpdatesCorrectSheet()
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
            { "errorMessage", "Sheet2 error" }
        });

        _handler.Execute(context, parameters);

        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].Validations[0].ErrorMessage));
        Assert.Equal("Sheet2 error", workbook.Worksheets[1].Validations[0].ErrorMessage);
    }

    #endregion

    #region Basic Set Messages Operations

    [Fact]
    public void Execute_SetsErrorMessage()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "errorMessage", "Invalid value!" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Updated", result.Message);
        Assert.Contains("ErrorMessage=Invalid value!", result.Message);
        Assert.Equal("Invalid value!", workbook.Worksheets[0].Validations[0].ErrorMessage);
        Assert.True(workbook.Worksheets[0].Validations[0].ShowError);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsInputMessage()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "inputMessage", "Select a value" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("InputMessage=Select a value", result.Message);
        Assert.Equal("Select a value", workbook.Worksheets[0].Validations[0].InputMessage);
        Assert.True(workbook.Worksheets[0].Validations[0].ShowInput);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsBothMessages()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "errorMessage", "Error!" },
            { "inputMessage", "Help text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("ErrorMessage=Error!", result.Message);
        Assert.Contains("InputMessage=Help text", result.Message);
        Assert.Equal("Error!", workbook.Worksheets[0].Validations[0].ErrorMessage);
        Assert.Equal("Help text", workbook.Worksheets[0].Validations[0].InputMessage);
    }

    [Fact]
    public void Execute_EmptyErrorMessage_DisablesShowError()
    {
        var workbook = CreateWorkbookWithValidation();
        workbook.Worksheets[0].Validations[0].ErrorMessage = "Old error";
        workbook.Worksheets[0].Validations[0].ShowError = true;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "errorMessage", "" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Validations[0].ErrorMessage);
        Assert.False(workbook.Worksheets[0].Validations[0].ShowError);
    }

    [Fact]
    public void Execute_EmptyInputMessage_DisablesShowInput()
    {
        var workbook = CreateWorkbookWithValidation();
        workbook.Worksheets[0].Validations[0].InputMessage = "Old help";
        workbook.Worksheets[0].Validations[0].ShowInput = true;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "inputMessage", "" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Validations[0].InputMessage);
        Assert.False(workbook.Worksheets[0].Validations[0].ShowInput);
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
    public void Execute_WithNoMessages_ReturnsNoChanges()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("No changes", result.Message);
    }

    [Fact]
    public void Execute_ReturnsValidationIndex()
    {
        var workbook = CreateWorkbookWithValidation();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "validationIndex", 0 },
            { "errorMessage", "Error" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithValidation()
    {
        var workbook = new Workbook();
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Opt1,Opt2,Opt3");
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
