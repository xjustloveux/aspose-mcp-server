using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.DataValidation;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsEmptyListForNoValidations()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Empty(json.RootElement.GetProperty("items").EnumerateArray());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];
        Assert.Equal("List", firstItem.GetProperty("type").GetString());
        Assert.Equal("Red,Green,Blue", firstItem.GetProperty("formula1").GetString());
        Assert.Equal("Select a color", firstItem.GetProperty("errorMessage").GetString());
        Assert.True(firstItem.GetProperty("showError").GetBoolean());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("Sheet2", json.RootElement.GetProperty("worksheetName").GetString());
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        AddValidationToSheet(workbook.Worksheets[0], "A1", "Yes,No");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("Sheet1", json.RootElement.GetProperty("worksheetName").GetString());
    }

    #endregion

    #region JSON Structure

    [Fact]
    public void Execute_ReturnsValidJsonStructure()
    {
        var workbook = CreateWorkbookWithValidations(1);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
        Assert.True(json.RootElement.TryGetProperty("items", out _));
    }

    [Fact]
    public void Execute_ReturnsValidationIndex()
    {
        var workbook = CreateWorkbookWithValidations(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var index = 0;
        foreach (var item in items.EnumerateArray())
        {
            Assert.Equal(index, item.GetProperty("index").GetInt32());
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
