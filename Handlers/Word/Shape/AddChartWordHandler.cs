using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using ImageType = Aspose.Cells.Drawing.ImageType;
using IOFile = System.IO.File;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding charts to Word documents.
/// </summary>
public class AddChartWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_chart";

    /// <summary>
    ///     Adds a chart to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: data
    ///     Optional: chartType, chartTitle, chartWidth, chartHeight, paragraphIndex, alignment
    /// </param>
    /// <returns>Success message with chart details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var chartParams = ExtractChartParameters(parameters);
        var tableData = ParseChartData(chartParams.Data);

        var doc = context.Document;
        var tempExcelPath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.xlsx");
        try
        {
            var tempImagePath = CreateChartImage(tableData, chartParams, tempExcelPath);
            InsertChartInDocument(doc, tempImagePath, chartParams);
            CleanupTempFile(tempImagePath);

            MarkModified(context);
            return $"Successfully added chart. Type: {chartParams.ChartType}, Data rows: {tableData.Count}.";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error creating chart: {ex.Message}", ex);
        }
        finally
        {
            CleanupTempFile(tempExcelPath);
        }
    }

    /// <summary>
    ///     Extracts chart parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted chart parameters.</returns>
    private static ChartParameters ExtractChartParameters(OperationParameters parameters)
    {
        return new ChartParameters(
            parameters.GetOptional("chartType", "column"),
            parameters.GetOptional<string[][]?>("data"),
            parameters.GetOptional<string?>("chartTitle"),
            parameters.GetOptional("chartWidth", 432.0),
            parameters.GetOptional("chartHeight", 252.0),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("alignment", "left")
        );
    }

    /// <summary>
    ///     Parses chart data from 2D array.
    /// </summary>
    /// <param name="data">The chart data array.</param>
    /// <returns>The parsed data as list of lists.</returns>
    /// <exception cref="ArgumentException">Thrown when data is empty.</exception>
    private static List<List<string>> ParseChartData(string[][]? data)
    {
        if (data == null || data.Length == 0)
            throw new ArgumentException("Chart data cannot be empty");

        return data.Select(row => row.ToList()).ToList();
    }

    /// <summary>
    ///     Creates a chart image from the data.
    /// </summary>
    /// <param name="tableData">The table data.</param>
    /// <param name="chartParams">The chart parameters.</param>
    /// <param name="tempExcelPath">The temporary Excel file path.</param>
    /// <returns>The path to the created chart image.</returns>
    private static string CreateChartImage(List<List<string>> tableData, ChartParameters chartParams,
        string tempExcelPath)
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        PopulateWorksheet(worksheet, tableData);
        var chart = CreateChart(worksheet, tableData, chartParams);
        workbook.Save(tempExcelPath);

        var tempImagePath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.png");
        chart.ToImage(tempImagePath, ImageType.Png);
        return tempImagePath;
    }

    /// <summary>
    ///     Populates an Excel worksheet with the table data.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="tableData">The table data.</param>
    private static void PopulateWorksheet(Worksheet worksheet, List<List<string>> tableData)
    {
        for (var i = 0; i < tableData.Count; i++)
        for (var j = 0; j < tableData[i].Count; j++)
        {
            var cellValue = tableData[i][j];
            if (double.TryParse(cellValue, out var numValue) && i > 0)
                worksheet.Cells[i, j].PutValue(numValue);
            else
                worksheet.Cells[i, j].PutValue(cellValue);
        }
    }

    /// <summary>
    ///     Creates a chart in the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="tableData">The table data.</param>
    /// <param name="chartParams">The chart parameters.</param>
    /// <returns>The created chart.</returns>
    private static Chart CreateChart(Worksheet worksheet, List<List<string>> tableData, ChartParameters chartParams)
    {
        var maxCol = tableData.Max(r => r.Count);
        var dataRange = $"A1:{Convert.ToChar(64 + maxCol)}{tableData.Count}";
        var chartTypeEnum = ParseChartType(chartParams.ChartType);

        var chartIndex = worksheet.Charts.Add(chartTypeEnum, 0, tableData.Count + 2, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.SetChartDataRange(dataRange, true);

        if (!string.IsNullOrEmpty(chartParams.ChartTitle))
            chart.Title.Text = chartParams.ChartTitle;

        return chart;
    }

    /// <summary>
    ///     Parses chart type string to ChartType enum.
    /// </summary>
    /// <param name="chartType">The chart type string.</param>
    /// <returns>The corresponding ChartType enum value.</returns>
    private static ChartType ParseChartType(string chartType)
    {
        return chartType.ToLower() switch
        {
            "bar" => ChartType.Bar,
            "line" => ChartType.Line,
            "pie" => ChartType.Pie,
            "area" => ChartType.Area,
            "scatter" => ChartType.Scatter,
            "doughnut" => ChartType.Doughnut,
            _ => ChartType.Column
        };
    }

    /// <summary>
    ///     Inserts the chart image into the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="imagePath">The chart image path.</param>
    /// <param name="chartParams">The chart parameters.</param>
    private static void InsertChartInDocument(Document doc, string imagePath, ChartParameters chartParams)
    {
        var builder = new DocumentBuilder(doc);
        MoveToInsertPosition(builder, doc, chartParams.ParagraphIndex);
        builder.ParagraphFormat.Alignment = WordShapeHelper.ParseAlignment(chartParams.Alignment);

        var shape = builder.InsertImage(imagePath);
        shape.Width = chartParams.ChartWidth;
        shape.Height = chartParams.ChartHeight;
        shape.WrapType = WrapType.Inline;
    }

    /// <summary>
    ///     Moves the builder to the insert position.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    private static void MoveToInsertPosition(DocumentBuilder builder, Document doc, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphIndex.Value == -1)
        {
            if (paragraphs.Count > 0 && paragraphs[0] is WordParagraph firstPara)
                builder.MoveTo(firstPara);
            return;
        }

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} out of range (total paragraphs: {paragraphs.Count})");

        if (paragraphs[paragraphIndex.Value] is WordParagraph targetPara)
            builder.MoveTo(targetPara);
        else
            throw new ArgumentException($"Cannot find paragraph at index {paragraphIndex.Value}");
    }

    /// <summary>
    ///     Cleans up a temporary file.
    /// </summary>
    /// <param name="filePath">The file path to delete.</param>
    private static void CleanupTempFile(string filePath)
    {
        if (!IOFile.Exists(filePath)) return;
        try
        {
            IOFile.Delete(filePath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error deleting temp file: {ex.Message}");
        }
    }

    /// <summary>
    ///     Record to hold chart creation parameters.
    /// </summary>
    private sealed record ChartParameters(
        string ChartType,
        string[][]? Data,
        string? ChartTitle,
        double ChartWidth,
        double ChartHeight,
        int? ParagraphIndex,
        string Alignment);
}
