using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;
using WordShape = Aspose.Words.Drawing.Shape;
using IOFile = System.IO.File;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetFooterImageHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_footer_image";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var alignment = parameters.GetOptional("alignment", "left");
        var imageWidth = parameters.GetOptional<double?>("imageWidth");
        var imageHeight = parameters.GetOptional<double?>("imageHeight");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var headerFooterType = parameters.GetOptional("headerFooterType", "primary");
        var isFloating = parameters.GetOptional("isFloating", false);
        var removeExisting = parameters.GetOptional("removeExisting", true);

        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath cannot be null or empty");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!IOFile.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (removeExisting)
            {
                var existingShapes = footer.GetChildNodes(NodeType.Shape, true).Cast<WordShape>()
                    .Where(s => s.HasImage).ToList();
                foreach (var existingShape in existingShapes) existingShape.Remove();
            }

            var footerPara = new WordParagraph(doc);
            footer.AppendChild(footerPara);

            var builder = new DocumentBuilder(doc);
            builder.MoveTo(footerPara);

            var paraAlignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                _ => ParagraphAlignment.Left
            };
            builder.ParagraphFormat.Alignment = paraAlignment;

            var shape = builder.InsertImage(imagePath);
            if (imageWidth.HasValue) shape.Width = imageWidth.Value;
            if (imageHeight.HasValue) shape.Height = imageHeight.Value;

            if (isFloating)
            {
                shape.WrapType = WrapType.Square;
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.BottomMargin;

                var pageWidth = section.PageSetup.PageWidth;
                var leftMargin = section.PageSetup.LeftMargin;
                var rightMargin = section.PageSetup.RightMargin;

                switch (alignment.ToLower())
                {
                    case "center":
                        shape.Left = (pageWidth - shape.Width) / 2;
                        break;
                    case "right":
                        shape.Left = pageWidth - rightMargin - shape.Width;
                        break;
                    default:
                        shape.Left = leftMargin;
                        break;
                }

                shape.Top = 0;
            }
            else
            {
                footerPara.ParagraphFormat.Alignment = paraAlignment;
            }
        }

        MarkModified(context);

        var floatingDesc = isFloating ? " (floating)" : "";
        return Success($"Footer image set{floatingDesc}");
    }
}
