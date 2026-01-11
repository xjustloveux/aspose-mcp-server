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
        var chartType = parameters.GetOptional("chartType", "column");
        var data = parameters.GetOptional<string[][]?>("data");
        var chartTitle = parameters.GetOptional<string?>("chartTitle");
        var chartWidth = parameters.GetOptional("chartWidth", 432.0);
        var chartHeight = parameters.GetOptional("chartHeight", 252.0);
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var alignment = parameters.GetOptional("alignment", "left");

        if (data == null || data.Length == 0)
            throw new ArgumentException("Chart data cannot be empty");

        List<List<string>> tableData = [];
        foreach (var row in data)
        {
            List<string> rowData = [];
            foreach (var cell in row)
                rowData.Add(cell);
            tableData.Add(rowData);
        }

        if (tableData.Count == 0)
            throw new ArgumentException("Cannot parse chart data");

        var doc = context.Document;
        var tempExcelPath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.xlsx");
        try
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            for (var i = 0; i < tableData.Count; i++)
            for (var j = 0; j < tableData[i].Count; j++)
            {
                var cellValue = tableData[i][j];
                if (double.TryParse(cellValue, out var numValue) && i > 0)
                    worksheet.Cells[i, j].PutValue(numValue);
                else
                    worksheet.Cells[i, j].PutValue(cellValue);
            }

            var maxCol = tableData.Max(r => r.Count);
            var dataRange = $"A1:{Convert.ToChar(64 + maxCol)}{tableData.Count}";

            var chartTypeEnum = chartType.ToLower() switch
            {
                "bar" => ChartType.Bar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.Scatter,
                "doughnut" => ChartType.Doughnut,
                _ => ChartType.Column
            };

            var chartIndex = worksheet.Charts.Add(chartTypeEnum, 0, tableData.Count + 2, 20, 10);
            var chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange(dataRange, true);

            if (!string.IsNullOrEmpty(chartTitle))
                chart.Title.Text = chartTitle;

            workbook.Save(tempExcelPath);

            var tempImagePath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.png");
            chart.ToImage(tempImagePath, ImageType.Png);

            var builder = new DocumentBuilder(doc);

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    if (paragraphs.Count > 0)
                        if (paragraphs[0] is WordParagraph firstPara)
                            builder.MoveTo(firstPara);
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    if (paragraphs[paragraphIndex.Value] is WordParagraph targetPara)
                        builder.MoveTo(targetPara);
                    else
                        throw new ArgumentException($"Cannot find paragraph at index {paragraphIndex.Value}");
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} out of range (total paragraphs: {paragraphs.Count})");
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
            }

            builder.ParagraphFormat.Alignment = WordShapeHelper.ParseAlignment(alignment);

            var shape = builder.InsertImage(tempImagePath);
            shape.Width = chartWidth;
            shape.Height = chartHeight;
            shape.WrapType = WrapType.Inline;

            if (IOFile.Exists(tempImagePath))
                try
                {
                    IOFile.Delete(tempImagePath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Error deleting temp image file: {ex.Message}");
                }

            MarkModified(context);

            return $"Successfully added chart. Type: {chartType}, Data rows: {tableData.Count}.";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error creating chart: {ex.Message}", ex);
        }
        finally
        {
            if (IOFile.Exists(tempExcelPath))
                try
                {
                    IOFile.Delete(tempExcelPath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[WARN] Error deleting temp Excel file: {ex.Message}");
                }
        }
    }
}
