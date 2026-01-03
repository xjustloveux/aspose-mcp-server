using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelConditionalFormattingToolTests : ExcelTestBase
{
    private readonly ExcelConditionalFormattingTool _tool;

    public ExcelConditionalFormattingToolTests()
    {
        _tool = new ExcelConditionalFormattingTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddConditionalFormatting_ShouldAddFormatting()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_conditional_formatting.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_conditional_formatting_output.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A5",
            condition: "GreaterThan",
            value: "10",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.ConditionalFormattings.Count > 0, "Conditional formatting should be added");
    }

    [Fact]
    public void GetConditionalFormatting_ShouldReturnFormatting()
    {
        var workbookPath = CreateExcelWorkbook("test_get_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);
        var result = _tool.Execute(
            "get",
            workbookPath,
            conditionalFormattingIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void DeleteConditionalFormatting_ShouldDeleteFormatting()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var formatCountBefore = worksheet.ConditionalFormattings.Count;
        Assert.True(formatCountBefore > 0, "Conditional formatting should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_conditional_formatting_output.xlsx");
        _tool.Execute(
            "delete",
            workbookPath,
            conditionalFormattingIndex: 0,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var formatCountAfter = resultWorksheet.ConditionalFormattings.Count;
        Assert.True(formatCountAfter < formatCountBefore,
            $"Conditional formatting should be deleted. Before: {formatCountBefore}, After: {formatCountAfter}");
    }

    [Fact]
    public void EditConditionalFormatting_ShouldEditFormatting()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_conditional_formatting_output.xlsx");
        _tool.Execute(
            "edit",
            workbookPath,
            conditionalFormattingIndex: 0,
            conditionIndex: 0,
            condition: "LessThan",
            value: "20",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.ConditionalFormattings.Count > 0,
            "Conditional formatting should exist after editing");
    }

    [Fact]
    public void Add_WithInvalidRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_range.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            range: "InvalidRange",
            condition: "GreaterThan",
            value: "10",
            outputPath: outputPath));
        Assert.Contains("Invalid range format", ex.Message);
    }

    [Fact]
    public void Add_WithBetweenCondition_ShouldUseBothFormulas()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_between_condition.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_between_condition_output.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A5",
            condition: "Between",
            value: "10",
            formula2: "50",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.ConditionalFormattings.Count > 0, "Conditional formatting should be added");
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid_index.xlsx");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            conditionalFormattingIndex: 999,
            outputPath: outputPath));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithNoFormattings_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_no_formatting.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No conditional formattings found", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", workbookPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var outputPath = CreateTestFilePath("test_missing_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, condition: "GreaterThan", value: "10", outputPath: outputPath));

        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingCondition_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_condition.xlsx");
        var outputPath = CreateTestFilePath("test_missing_condition_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A5", value: "10", outputPath: outputPath));

        Assert.Contains("condition is required", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_cf.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add_cf.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, range: "A1:A5", condition: "GreaterThan", value: "10");
        Assert.Contains("Conditional formatting added", result);

        // Verify in-memory workbook has the formatting
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(sessionWorkbook.Worksheets[0].ConditionalFormattings.Count > 0,
            "Conditional formatting should be added in memory");
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit_cf.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, conditionalFormattingIndex: 0, conditionIndex: 0,
            condition: "LessThan", value: "5");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(sessionWorkbook.Worksheets[0].ConditionalFormattings.Count > 0,
            "Conditional formatting should exist in memory after edit");
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete_cf.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);

        // Verify formatting exists before delete
        var beforeWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(beforeWorkbook.Worksheets[0].ConditionalFormattings.Count > 0);
        _tool.Execute("delete", sessionId: sessionId, conditionalFormattingIndex: 0);

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(sessionWorkbook.Worksheets[0].ConditionalFormattings);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}