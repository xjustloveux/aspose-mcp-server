using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class RenameExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly RenameExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Rename()
    {
        Assert.Equal("rename", _handler.Operation);
    }

    #endregion

    #region Preserve Other Sheets

    [Fact]
    public void Execute_PreservesOtherSheets()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "newName", "Renamed" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sheet1", workbook.Worksheets[0].Name);
        Assert.Equal("Renamed", workbook.Worksheets[1].Name);
        Assert.Equal("Sheet3", workbook.Worksheets[2].Name);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsOldAndNewNameInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "newName", "NewName" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet1", result);
        Assert.Contains("NewName", result);
    }

    #endregion

    #region Basic Rename Operations

    [Fact]
    public void Execute_RenamesSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "newName", "RenamedSheet" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("renamed", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("RenamedSheet", workbook.Worksheets[0].Name);
        AssertModified(context);
    }

    [Theory]
    [InlineData("NewName")]
    [InlineData("Data")]
    [InlineData("Report 2024")]
    public void Execute_RenamesWithVariousNames(string newName)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "newName", newName }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(newName, workbook.Worksheets[0].Name);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_RenamesSheetAtVariousIndices(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", sheetIndex },
            { "newName", "Renamed" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Renamed", workbook.Worksheets[sheetIndex].Name);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "newName", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutNewName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("newName", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithDuplicateName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "newName", "Sheet2" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already exists", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", invalidIndex },
            { "newName", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
