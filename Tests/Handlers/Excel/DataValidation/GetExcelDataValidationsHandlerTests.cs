using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Results.Excel.DataValidation;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataValidation;

public class GetExcelDataValidationsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelDataValidationsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var workbook = CreateWorkbookWithValidations(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
        Assert.Equal(2, workbook.Worksheets[0].Validations.Count);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsValidations()
    {
        var workbook = CreateWorkbookWithValidations(2);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        Assert.Equal(2, result.Count);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsEmptyListForNoValidations()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
    }

    [Fact]
    public void Execute_ReturnsValidationDetails()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 };
        var index = sheet.Validations.Add(area);
        var validation = sheet.Validations[index];
        validation.Type = ValidationType.List;
        validation.Formula1 = "Red,Green,Blue";
        validation.ErrorMessage = "Select a color";
        validation.ShowError = true;
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        var firstItem = result.Items[0];
        Assert.Equal("List", firstItem.Type);
        Assert.Equal("Red,Green,Blue", firstItem.Formula1);
        Assert.Equal("Select a color", firstItem.ErrorMessage);
        Assert.True(firstItem.ShowError);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Yes,No");
        AddValidationToSheet(workbook.Worksheets[1], "B1", "X,Y,Z");
        AddValidationToSheet(workbook.Worksheets[1], "B2", "A,B,C");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        Assert.Equal(2, result.Count);
        Assert.Equal("Sheet2", result.WorksheetName);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Yes,No");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        Assert.Equal(1, result.Count);
        Assert.Equal("Sheet1", result.WorksheetName);
    }

    #endregion

    #region JSON Structure

    [Fact]
    public void Execute_ReturnsValidJsonStructure()
    {
        var workbook = CreateWorkbookWithValidations(1);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.WorksheetName);
        Assert.NotNull(result.Items);
    }

    [Fact]
    public void Execute_ReturnsValidationIndex()
    {
        var workbook = CreateWorkbookWithValidations(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetDataValidationsResult>(res);

        var index = 0;
        foreach (var item in result.Items)
        {
            Assert.Equal(index, item.Index);
            index++;
        }
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
