using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddSectionBreakTool : IAsposeTool
{
    public string Description => "Add a section break to a Word document with custom section settings";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            breakType = new
            {
                type = "string",
                description = "Section break type: nextPage, continuous, evenPage, oddPage (default: nextPage)",
                @enum = new[] { "nextPage", "continuous", "evenPage", "oddPage" }
            },
            orientation = new
            {
                type = "string",
                description = "Page orientation for new section: portrait, landscape (optional)",
                @enum = new[] { "portrait", "landscape" }
            },
            pageWidth = new
            {
                type = "number",
                description = "Page width in points (optional, e.g., 595 for A4)"
            },
            pageHeight = new
            {
                type = "number",
                description = "Page height in points (optional, e.g., 842 for A4)"
            },
            marginTop = new
            {
                type = "number",
                description = "Top margin in points (optional)"
            },
            marginBottom = new
            {
                type = "number",
                description = "Bottom margin in points (optional)"
            },
            marginLeft = new
            {
                type = "number",
                description = "Left margin in points (optional)"
            },
            marginRight = new
            {
                type = "number",
                description = "Right margin in points (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var breakType = arguments?["breakType"]?.GetValue<string>() ?? "nextPage";
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var pageWidth = arguments?["pageWidth"]?.GetValue<double?>();
        var pageHeight = arguments?["pageHeight"]?.GetValue<double?>();
        var marginTop = arguments?["marginTop"]?.GetValue<double?>();
        var marginBottom = arguments?["marginBottom"]?.GetValue<double?>();
        var marginLeft = arguments?["marginLeft"]?.GetValue<double?>();
        var marginRight = arguments?["marginRight"]?.GetValue<double?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Move to end of document
        builder.MoveToDocumentEnd();
        
        // Insert section break
        var sectionStart = breakType.ToLower() switch
        {
            "continuous" => SectionStart.Continuous,
            "evenpage" => SectionStart.EvenPage,
            "oddpage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };
        
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Get the new section
        var section = builder.CurrentSection;
        section.PageSetup.SectionStart = sectionStart;

        // Apply page settings
        if (!string.IsNullOrEmpty(orientation))
        {
            section.PageSetup.Orientation = orientation.ToLower() == "landscape" 
                ? Orientation.Landscape 
                : Orientation.Portrait;
        }

        if (pageWidth.HasValue)
            section.PageSetup.PageWidth = pageWidth.Value;

        if (pageHeight.HasValue)
            section.PageSetup.PageHeight = pageHeight.Value;

        if (marginTop.HasValue)
            section.PageSetup.TopMargin = marginTop.Value;

        if (marginBottom.HasValue)
            section.PageSetup.BottomMargin = marginBottom.Value;

        if (marginLeft.HasValue)
            section.PageSetup.LeftMargin = marginLeft.Value;

        if (marginRight.HasValue)
            section.PageSetup.RightMargin = marginRight.Value;

        doc.Save(outputPath);

        var result = $"成功添加分節符號\n";
        result += $"類型: {breakType}\n";
        if (!string.IsNullOrEmpty(orientation)) result += $"方向: {orientation}\n";
        if (pageWidth.HasValue || pageHeight.HasValue) 
            result += $"頁面尺寸: {pageWidth ?? section.PageSetup.PageWidth} x {pageHeight ?? section.PageSetup.PageHeight} pt\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

