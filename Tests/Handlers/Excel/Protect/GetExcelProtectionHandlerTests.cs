using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

public class GetExcelProtectionHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelProtectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Specific Sheet

    [Fact]
    public void Execute_WithSheetIndex_ReturnsOnlySpecifiedSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("totalWorksheets").GetInt32());
        var worksheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.Equal(1, worksheet.GetProperty("index").GetInt32());
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
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsProtectionInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("totalWorksheets", out _));
        Assert.True(json.RootElement.TryGetProperty("worksheets", out _));
    }

    [Fact]
    public void Execute_ReturnsAllWorksheets()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("totalWorksheets").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("worksheets").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsProtectionStatus()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var worksheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.True(worksheet.TryGetProperty("isProtected", out _));
        Assert.True(worksheet.TryGetProperty("name", out _));
        Assert.True(worksheet.TryGetProperty("index", out _));
    }

    [Fact]
    public void Execute_ReturnsProtectionDetails()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var worksheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.True(worksheet.TryGetProperty("allowSelectingLockedCell", out _));
        Assert.True(worksheet.TryGetProperty("allowSelectingUnlockedCell", out _));
        Assert.True(worksheet.TryGetProperty("allowFormattingCell", out _));
        Assert.True(worksheet.TryGetProperty("allowFiltering", out _));
        Assert.True(worksheet.TryGetProperty("allowSorting", out _));
    }

    [Fact]
    public void Execute_WithProtectedSheet_ReturnsIsProtectedTrue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);

        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var worksheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.True(worksheet.GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public void Execute_WithUnprotectedSheet_ReturnsIsProtectedFalse()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var worksheet = json.RootElement.GetProperty("worksheets")[0];
        Assert.False(worksheet.GetProperty("isProtected").GetBoolean());
    }

    #endregion
}
