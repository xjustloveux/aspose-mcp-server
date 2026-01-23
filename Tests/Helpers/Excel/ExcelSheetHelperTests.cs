using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelSheetHelperTests
{
    #region ValidateSheetName Tests - Valid Names

    [Theory]
    [InlineData("Sheet1")]
    [InlineData("MySheet")]
    [InlineData("Data 2024")]
    [InlineData("Summary-Report")]
    [InlineData("Sheet_With_Underscores")]
    [InlineData("Sheet.With.Dots")]
    [InlineData("Sheet(With)Parentheses")]
    [InlineData("Sheet'With'Quotes")]
    [InlineData("123")]
    [InlineData("A")]
    public void ValidateSheetName_WithValidNames_DoesNotThrow(string name)
    {
        var exception = Record.Exception(() =>
            ExcelSheetHelper.ValidateSheetName(name, "sheetName"));

        Assert.Null(exception);
    }

    #endregion

    #region ValidateSheetName Tests - Empty/Whitespace

    [Fact]
    public void ValidateSheetName_WithNull_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName(null!, "sheetName"));

        Assert.Contains("cannot be empty", ex.Message);
    }

    [Fact]
    public void ValidateSheetName_WithEmptyString_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName("", "sheetName"));

        Assert.Contains("cannot be empty", ex.Message);
    }

    [Fact]
    public void ValidateSheetName_WithWhitespace_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName("   ", "sheetName"));

        Assert.Contains("cannot be empty", ex.Message);
    }

    #endregion

    #region ValidateSheetName Tests - Length

    [Fact]
    public void ValidateSheetName_WithExactly31Characters_DoesNotThrow()
    {
        var name = new string('A', 31);

        var exception = Record.Exception(() =>
            ExcelSheetHelper.ValidateSheetName(name, "sheetName"));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSheetName_With32Characters_ThrowsArgumentException()
    {
        var name = new string('A', 32);

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName(name, "sheetName"));

        Assert.Contains("exceeds Excel's limit of 31 characters", ex.Message);
        Assert.Contains("length: 32", ex.Message);
    }

    [Fact]
    public void ValidateSheetName_With100Characters_ThrowsArgumentException()
    {
        var name = new string('X', 100);

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName(name, "sheetName"));

        Assert.Contains("exceeds Excel's limit of 31 characters", ex.Message);
    }

    #endregion

    #region ValidateSheetName Tests - Invalid Characters

    [Theory]
    [InlineData("Sheet\\Name", '\\')]
    [InlineData("Sheet/Name", '/')]
    [InlineData("Sheet?Name", '?')]
    [InlineData("Sheet*Name", '*')]
    [InlineData("Sheet[Name", '[')]
    [InlineData("Sheet]Name", ']')]
    [InlineData("Sheet:Name", ':')]
    public void ValidateSheetName_WithInvalidCharacter_ThrowsArgumentException(string name, char invalidChar)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName(name, "sheetName"));

        Assert.Contains($"contains invalid character '{invalidChar}'", ex.Message);
        Assert.Contains("Sheet names cannot contain: \\ / ? * [ ] :", ex.Message);
    }

    [Fact]
    public void ValidateSheetName_WithMultipleInvalidCharacters_ReportsFirstOne()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName("A\\B?C", "sheetName"));

        Assert.Contains("contains invalid character '\\'", ex.Message);
    }

    #endregion

    #region ValidateSheetName Tests - ParamName in Error Message

    [Fact]
    public void ValidateSheetName_WithCustomParamName_IncludesParamNameInMessage()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName("", "customParamName"));

        Assert.Contains("customParamName", ex.Message);
    }

    [Fact]
    public void ValidateSheetName_LengthError_IncludesActualName()
    {
        var name = new string('B', 35);

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelSheetHelper.ValidateSheetName(name, "targetSheet"));

        Assert.Contains(name, ex.Message);
        Assert.Contains("targetSheet", ex.Message);
    }

    #endregion
}
