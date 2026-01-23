using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;
using WordShape = Aspose.Words.Drawing.Shape;
using IOFile = System.IO.File;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Base class for header/footer image handlers.
///     Provides common functionality for setting images in headers and footers.
/// </summary>
public abstract class HeaderFooterImageHandlerBase : OperationHandlerBase<Document>
{
    /// <summary>
    ///     Whether this handler targets the header (true) or footer (false).
    /// </summary>
    protected abstract bool IsHeader { get; }

    /// <summary>
    ///     The display name for the target ("Header" or "Footer").
    /// </summary>
    protected abstract string TargetName { get; }

    /// <summary>
    ///     The vertical position for floating images.
    /// </summary>
    protected abstract RelativeVerticalPosition VerticalPosition { get; }

    /// <summary>
    ///     Sets an image in the document header or footer.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imagePath
    ///     Optional: alignment, imageWidth, imageHeight, sectionIndex, headerFooterType, isFloating, removeExisting
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractImageParameters(parameters);
        ValidateImagePath(p.ImagePath);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, IsHeader);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
            InsertImage(section, doc, hfType, p);

        MarkModified(context);
        return new SuccessResult { Message = $"{TargetName} image set{(p.IsFloating ? " (floating)" : "")}" };
    }

    /// <summary>
    ///     Extracts image parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted image parameters.</returns>
    private static ImageParameters ExtractImageParameters(OperationParameters parameters)
    {
        return new ImageParameters(
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("alignment", "left"),
            parameters.GetOptional<double?>("imageWidth"),
            parameters.GetOptional<double?>("imageHeight"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary"),
            parameters.GetOptional("isFloating", false),
            parameters.GetOptional("removeExisting", true)
        );
    }

    /// <summary>
    ///     Validates the image path.
    /// </summary>
    /// <param name="imagePath">The image path to validate.</param>
    /// <exception cref="ArgumentException">Thrown when image path is null or empty.</exception>
    /// <exception cref="FileNotFoundException">Thrown when image file is not found.</exception>
    private static void ValidateImagePath(string? imagePath)
    {
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath cannot be null or empty");
        SecurityHelper.ValidateFilePath(imagePath, nameof(imagePath), true);
        if (!IOFile.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");
    }

    /// <summary>
    ///     Inserts an image into the header or footer.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="hfType">The header/footer type.</param>
    /// <param name="p">The image parameters.</param>
    private void InsertImage(Section section, Document doc, HeaderFooterType hfType, ImageParameters p)
    {
        var headerFooter = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

        if (p.RemoveExisting)
            RemoveExistingImages(headerFooter);

        var para = new WordParagraph(doc);
        headerFooter.AppendChild(para);

        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para);

        var paraAlignment = GetParagraphAlignment(p.Alignment);
        builder.ParagraphFormat.Alignment = paraAlignment;

        var shape = builder.InsertImage(p.ImagePath);
        ApplyImageDimensions(shape, p.ImageWidth, p.ImageHeight);

        if (p.IsFloating)
            ApplyFloatingPosition(shape, section, p.Alignment, VerticalPosition);
        else
            para.ParagraphFormat.Alignment = paraAlignment;
    }

    /// <summary>
    ///     Removes existing images from the header/footer.
    /// </summary>
    /// <param name="hf">The header/footer.</param>
    private static void RemoveExistingImages(Aspose.Words.HeaderFooter hf)
    {
        var existingShapes = hf.GetChildNodes(NodeType.Shape, true).Cast<WordShape>()
            .Where(s => s.HasImage).ToList();
        foreach (var existingShape in existingShapes) existingShape.Remove();
    }

    /// <summary>
    ///     Gets the paragraph alignment from alignment string.
    /// </summary>
    /// <param name="alignment">The alignment string.</param>
    /// <returns>The paragraph alignment.</returns>
    private static ParagraphAlignment GetParagraphAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

    /// <summary>
    ///     Applies dimensions to an image shape.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <param name="width">The optional width.</param>
    /// <param name="height">The optional height.</param>
    private static void ApplyImageDimensions(WordShape shape, double? width, double? height)
    {
        if (width.HasValue) shape.Width = width.Value;
        if (height.HasValue) shape.Height = height.Value;
    }

    /// <summary>
    ///     Applies floating position to an image shape.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <param name="section">The document section.</param>
    /// <param name="alignment">The alignment string.</param>
    /// <param name="verticalPosition">The relative vertical position.</param>
    private static void ApplyFloatingPosition(WordShape shape, Section section, string alignment,
        RelativeVerticalPosition verticalPosition)
    {
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = verticalPosition;

        var pageWidth = section.PageSetup.PageWidth;
        var leftMargin = section.PageSetup.LeftMargin;
        var rightMargin = section.PageSetup.RightMargin;

        shape.Left = alignment.ToLower() switch
        {
            "center" => (pageWidth - shape.Width) / 2,
            "right" => pageWidth - rightMargin - shape.Width,
            _ => leftMargin
        };
        shape.Top = 0;
    }

    /// <summary>
    ///     Record to hold image insertion parameters.
    /// </summary>
    protected sealed record ImageParameters(
        string? ImagePath,
        string Alignment,
        double? ImageWidth,
        double? ImageHeight,
        int SectionIndex,
        string HeaderFooterType,
        bool IsFloating,
        bool RemoveExisting);
}
