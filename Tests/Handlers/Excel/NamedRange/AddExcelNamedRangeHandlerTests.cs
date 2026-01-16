using AsposeMcpServer.Handlers.Excel.NamedRange;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.NamedRange;

public class AddExcelNamedRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelNamedRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Cross-Sheet Range Reference

    [Fact]
    public void Execute_WithSheetReferenceInRange_ParsesCorrectly()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("DataSheet");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "CrossSheetRange" },
            { "range", "DataSheet!A1:D10" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("DataSheet", workbook.Worksheets.Names["CrossSheetRange"].RefersTo);
        AssertModified(context);
    }

    #endregion

    #region Various Range Formats

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:A1")]
    [InlineData("A1:Z100")]
    [InlineData("AA1:AZ50")]
    public void Execute_WithVariousRangeFormats_CreatesNamedRange(string range)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", $"Range_{range.Replace(":", "_")}" },
            { "range", range }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsNamedRange()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestRange" },
            { "range", "A1:B5" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithComment_SetsComment()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestRange" },
            { "range", "A1:B5" },
            { "comment", "Test comment" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Equal("Test comment", workbook.Worksheets.Names["TestRange"].Comment);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestRange" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithDuplicateName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:B5").Name = "ExistingRange";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "ExistingRange" },
            { "range", "C1:D5" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Sheet Index Tests

    [Fact]
    public void Execute_WithSheetIndex_AddsToSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "RangeOnSheet2" },
            { "range", "A1:C3" },
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("Sheet2", workbook.Worksheets.Names["RangeOnSheet2"].RefersTo);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestRange" },
            { "range", "A1:B5" },
            { "sheetIndex", 99 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
