using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding charts to Word documents using native Aspose.Words chart API.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var chartParams = ExtractChartParameters(parameters);
        var tableData = ParseChartData(chartParams.Data);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        MoveToInsertPosition(builder, doc, chartParams.ParagraphIndex);
        builder.ParagraphFormat.Alignment = WordShapeHelper.ParseAlignment(chartParams.Alignment);

        var chartType = ParseChartType(chartParams.ChartType);
        var shape = builder.InsertChart(chartType, chartParams.ChartWidth, chartParams.ChartHeight);
        var chart = shape.Chart;

        if (!string.IsNullOrEmpty(chartParams.ChartTitle))
            chart.Title.Text = chartParams.ChartTitle;

        chart.Series.Clear();
        PopulateChartSeries(chart, tableData, chartType);

        MarkModified(context);
        return new SuccessResult
            { Message = $"Successfully added chart. Type: {chartParams.ChartType}, Data rows: {tableData.Count}." };
    }

    /// <summary>
    ///     Populates chart series from the parsed table data.
    /// </summary>
    /// <param name="chart">The chart to populate.</param>
    /// <param name="tableData">The table data (first row = headers, subsequent rows = data).</param>
    /// <param name="chartType">The chart type for scatter-specific handling.</param>
    private static void PopulateChartSeries(Chart chart, List<List<string>> tableData, ChartType chartType)
    {
        if (tableData.Count < 2 || tableData[0].Count < 2)
            return;

        var headers = tableData[0];
        var dataRows = tableData.Skip(1).ToList();

        if (chartType == ChartType.Scatter)
        {
            PopulateScatterSeries(chart, headers, dataRows);
            return;
        }

        var categories = dataRows.Select(r => r[0]).ToArray();

        for (var col = 1; col < headers.Count; col++)
        {
            var seriesName = headers[col];
            var values = dataRows.Select(r => col < r.Count ? ParseDouble(r[col]) : 0.0).ToArray();
            chart.Series.Add(seriesName, categories, values);
        }
    }

    /// <summary>
    ///     Populates scatter chart series with numeric X and Y values.
    /// </summary>
    /// <param name="chart">The chart to populate.</param>
    /// <param name="headers">The header row.</param>
    /// <param name="dataRows">The data rows.</param>
    private static void PopulateScatterSeries(Chart chart, List<string> headers, List<List<string>> dataRows)
    {
        var xValues = dataRows.Select(r => ParseDouble(r[0])).ToArray();

        for (var col = 1; col < headers.Count; col++)
        {
            var seriesName = headers[col];
            var yValues = dataRows.Select(r => col < r.Count ? ParseDouble(r[col]) : 0.0).ToArray();
            chart.Series.Add(seriesName, xValues, yValues);
        }
    }

    /// <summary>
    ///     Parses a string value to double, returning 0 if parsing fails.
    /// </summary>
    /// <param name="value">The string value to parse.</param>
    /// <returns>The parsed double value, or 0 if parsing fails.</returns>
    private static double ParseDouble(string value)
    {
        return double.TryParse(value, out var result) ? result : 0.0;
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
