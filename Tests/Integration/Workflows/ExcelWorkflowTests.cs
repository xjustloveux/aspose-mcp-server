using Aspose.Cells;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Integration.Workflows;

/// <summary>
///     Integration tests for Excel workbook workflows.
/// </summary>
[Trait("Category", "Integration")]
public class ExcelWorkflowTests : TestBase
{
    private readonly ExcelCellTool _cellTool;
    private readonly ExcelChartTool _chartTool;
    private readonly ExcelFormulaTool _formulaTool;
    private readonly ExcelPivotTableTool _pivotTableTool;
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelWorkflowTests" /> class.
    /// </summary>
    public ExcelWorkflowTests()
    {
        var config = new SessionConfig { Enabled = true };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _cellTool = new ExcelCellTool(_sessionManager);
        _formulaTool = new ExcelFormulaTool(_sessionManager);
        _pivotTableTool = new ExcelPivotTableTool(_sessionManager);
        _chartTool = new ExcelChartTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Open-Edit-Save Workflow Tests

    /// <summary>
    ///     Verifies the complete open, edit, and save workflow for Excel workbooks.
    /// </summary>
    [Fact]
    public void Excel_OpenEditSave_Workflow()
    {
        // Step 1: Create and open workbook
        var originalPath = CreateExcelDocument();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Edit workbook - edit cell value
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "B1", value: "Modified");

        // Step 3: Save workbook
        var outputPath = CreateTestFilePath("excel_workflow_output.xlsx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        // Step 4: Verify changes persisted
        using var savedWorkbook = new Workbook(outputPath);
        Assert.Equal("Modified", savedWorkbook.Worksheets[0].Cells["B1"].StringValue);

        // Step 5: Close session
        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Formula Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding and calculating formulas.
    /// </summary>
    [Fact]
    public void Excel_AddFormulaCalculate_Workflow()
    {
        // Step 1: Create workbook with data
        var originalPath = CreateExcelDocumentWithData();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add formula
        _formulaTool.Execute("add", sessionId: openData.SessionId, cell: "C1", formula: "=A1+B1");

        // Step 3: Calculate formulas
        _formulaTool.Execute("calculate", sessionId: openData.SessionId);

        // Step 4: Save and verify
        var outputPath = CreateTestFilePath("formula_workflow.xlsx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        using var savedWorkbook = new Workbook(outputPath);
        savedWorkbook.CalculateFormula();
        var result = savedWorkbook.Worksheets[0].Cells["C1"].Value;
        Assert.Equal(30, Convert.ToInt32(result));

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Import Data Workflow Tests

    /// <summary>
    ///     Verifies the workflow of importing data into cells.
    /// </summary>
    [Fact]
    public void Excel_ImportData_Workflow()
    {
        // Step 1: Create and open empty workbook
        var originalPath = CreateEmptyExcelDocument();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Edit multiple cell values
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "A1", value: "Product");
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "B1", value: "Price");
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "A2", value: "Apple");
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "B2", value: "1.50");
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "A3", value: "Banana");
        _cellTool.Execute("edit", sessionId: openData.SessionId, cell: "B3", value: "0.75");

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("import_data_workflow.xlsx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        using var savedWorkbook = new Workbook(outputPath);
        Assert.Equal("Product", savedWorkbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Apple", savedWorkbook.Worksheets[0].Cells["A2"].StringValue);
        Assert.Equal("Banana", savedWorkbook.Worksheets[0].Cells["A3"].StringValue);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region PivotTable Workflow Tests

    /// <summary>
    ///     Verifies the workflow of creating a pivot table.
    /// </summary>
    [Fact]
    public void Excel_CreatePivotTable_Workflow()
    {
        // Step 1: Create workbook with data for pivot table
        var originalPath = CreateExcelDocumentWithPivotData();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add pivot table
        _pivotTableTool.Execute("add",
            sessionId: openData.SessionId,
            sourceRange: "A1:C5",
            destCell: "E1");

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("pivot_table_workflow.xlsx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        using var savedWorkbook = new Workbook(outputPath);
        Assert.True(savedWorkbook.Worksheets[0].PivotTables.Count > 0);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Chart Workflow Tests

    /// <summary>
    ///     Verifies the workflow of creating a chart.
    /// </summary>
    [Fact]
    public void Excel_CreateChart_Workflow()
    {
        // Step 1: Create workbook with data for chart
        var originalPath = CreateExcelDocumentWithChartData();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add chart
        _chartTool.Execute("add",
            sessionId: openData.SessionId,
            chartType: "Column",
            dataRange: "A1:B4",
            title: "Sales Chart");

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("chart_workflow.xlsx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        using var savedWorkbook = new Workbook(outputPath);
        Assert.True(savedWorkbook.Worksheets[0].Charts.Count > 0);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Helper Methods

    private string CreateExcelDocument()
    {
        var path = CreateTestFilePath($"excel_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Original";
        workbook.Save(path);
        return path;
    }

    private string CreateEmptyExcelDocument()
    {
        var path = CreateTestFilePath($"excel_empty_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Save(path);
        return path;
    }

    private string CreateExcelDocumentWithData()
    {
        var path = CreateTestFilePath($"excel_data_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Value = 20;
        workbook.Save(path);
        return path;
    }

    private string CreateExcelDocumentWithPivotData()
    {
        var path = CreateTestFilePath($"excel_pivot_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        // Header row
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Product";
        cells["C1"].Value = "Sales";

        // Data rows
        cells["A2"].Value = "Fruit";
        cells["B2"].Value = "Apple";
        cells["C2"].Value = 100;

        cells["A3"].Value = "Fruit";
        cells["B3"].Value = "Banana";
        cells["C3"].Value = 150;

        cells["A4"].Value = "Vegetable";
        cells["B4"].Value = "Carrot";
        cells["C4"].Value = 80;

        cells["A5"].Value = "Vegetable";
        cells["B5"].Value = "Potato";
        cells["C5"].Value = 120;

        workbook.Save(path);
        return path;
    }

    private string CreateExcelDocumentWithChartData()
    {
        var path = CreateTestFilePath($"excel_chart_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        // Header row
        cells["A1"].Value = "Month";
        cells["B1"].Value = "Sales";

        // Data rows
        cells["A2"].Value = "Jan";
        cells["B2"].Value = 1000;

        cells["A3"].Value = "Feb";
        cells["B3"].Value = 1500;

        cells["A4"].Value = "Mar";
        cells["B4"].Value = 1200;

        workbook.Save(path);
        return path;
    }

    #endregion
}
