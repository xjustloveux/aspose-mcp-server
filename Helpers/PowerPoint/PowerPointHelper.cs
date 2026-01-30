using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.SmartArt;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Helper class for common PowerPoint operations to reduce code duplication
/// </summary>
public static class PowerPointHelper
{
    /// <summary>
    ///     Validates slide index and throws exception if invalid
    /// </summary>
    /// <param name="slideIndex">Slide index to validate</param>
    /// <param name="presentation">Presentation to check against</param>
    /// <exception cref="ArgumentException">Thrown if slide index is invalid</exception>
    public static void ValidateSlideIndex(int slideIndex, IPresentation presentation)
    {
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException(
                $"Slide index {slideIndex} is out of range (presentation has {presentation.Slides.Count} slides)");
    }

    /// <summary>
    ///     Gets a slide with validation
    /// </summary>
    /// <param name="presentation">Presentation to get slide from</param>
    /// <param name="slideIndex">Slide index</param>
    /// <returns>Slide</returns>
    /// <exception cref="ArgumentException">Thrown if slide index is invalid</exception>
    public static ISlide GetSlide(IPresentation presentation, int slideIndex)
    {
        ValidateSlideIndex(slideIndex, presentation);
        return presentation.Slides[slideIndex];
    }

    /// <summary>
    ///     Validates shape index and throws exception if invalid
    /// </summary>
    /// <param name="shapeIndex">Shape index to validate</param>
    /// <param name="slide">Slide to check against</param>
    /// <exception cref="ArgumentException">Thrown if shape index is invalid</exception>
    public static void ValidateShapeIndex(int shapeIndex, ISlide slide)
    {
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentException(
                $"Shape index {shapeIndex} is out of range (slide has {slide.Shapes.Count} shapes)");
    }

    /// <summary>
    ///     Gets a shape with validation
    /// </summary>
    /// <param name="slide">Slide to get shape from</param>
    /// <param name="shapeIndex">Shape index</param>
    /// <returns>Shape</returns>
    /// <exception cref="ArgumentException">Thrown if shape index is invalid</exception>
    public static IShape GetShape(ISlide slide, int shapeIndex)
    {
        ValidateShapeIndex(shapeIndex, slide);
        return slide.Shapes[shapeIndex];
    }

    /// <summary>
    ///     Validates collection index and throws exception if invalid
    /// </summary>
    /// <typeparam name="T">Collection item type</typeparam>
    /// <param name="index">Index to validate</param>
    /// <param name="collection">Collection to check against</param>
    /// <param name="itemName">Name of the item type for error message</param>
    /// <exception cref="ArgumentException">Thrown if index is invalid</exception>
    public static void ValidateCollectionIndex<T>(int index, ICollection<T> collection, string itemName = "Item")
    {
        if (index < 0 || index >= collection.Count)
            throw new ArgumentException(
                $"{itemName} index {index} is out of range (collection has {collection.Count} {itemName.ToLower()}s)");
    }

    /// <summary>
    ///     Validates collection index for collections with Count property (supports Aspose collections)
    /// </summary>
    /// <param name="index">Index to validate</param>
    /// <param name="count">Collection count</param>
    /// <param name="itemName">Name of the item type for error message</param>
    /// <exception cref="ArgumentException">Thrown if index is invalid</exception>
    public static void ValidateCollectionIndex(int index, int count, string itemName = "Item")
    {
        if (index < 0 || index >= count)
            throw new ArgumentException(
                $"{itemName} index {index} is out of range (collection has {count} {itemName.ToLower()}s)");
    }

    /// <summary>
    ///     Extracts text from a shape recursively, including tables, SmartArt, and group shapes.
    /// </summary>
    /// <param name="shape">The shape to extract text from.</param>
    /// <param name="textContent">List to add extracted text.</param>
    public static void
        ExtractTextFromShape(IShape shape,
            List<string> textContent)
    {
        switch (shape)
        {
            case IAutoShape { TextFrame.Text: var text } when !string.IsNullOrWhiteSpace(text):
                textContent.Add(text);
                break;
            case ITable table:
                foreach (var row in table.Rows)
                foreach (var cell in row.Where(cell => !string.IsNullOrWhiteSpace(cell.TextFrame?.Text)))
                    textContent.Add(cell.TextFrame.Text);

                break;
            case ISmartArt smartArt:
                foreach (var node in smartArt.AllNodes.Where(node => !string.IsNullOrWhiteSpace(node.TextFrame?.Text)))
                    textContent.Add(node.TextFrame.Text);
                break;
            case IGroupShape groupShape:
                foreach (var childShape in groupShape.Shapes)
                    ExtractTextFromShape(childShape, textContent);
                break;
        }
    }

    /// <summary>
    ///     Counts text characters in a shape recursively.
    /// </summary>
    /// <param name="shape">The shape to count characters from.</param>
    /// <returns>Total character count.</returns>
    public static int
        CountTextCharacters(IShape shape)
    {
        var count = 0;
        switch (shape)
        {
            case IAutoShape { TextFrame.Text: var text } when !string.IsNullOrWhiteSpace(text):
                count += text.Length;
                break;
            case ITable table:
                foreach (var row in table.Rows)
                foreach (var cell in row.Where(cell => !string.IsNullOrWhiteSpace(cell.TextFrame?.Text)))
                    count += cell.TextFrame.Text.Length;

                break;
            case ISmartArt smartArt:
                foreach (var node in smartArt.AllNodes.Where(node => !string.IsNullOrWhiteSpace(node.TextFrame?.Text)))
                    count += node.TextFrame.Text.Length;
                break;
            case IGroupShape groupShape:
                foreach (var childShape in groupShape.Shapes)
                    count += CountTextCharacters(childShape);
                break;
        }

        return count;
    }

    /// <summary>
    ///     Counts shape types for statistics.
    /// </summary>
    /// <param name="shape">The shape to categorize.</param>
    /// <param name="images">Reference to images count.</param>
    /// <param name="tables">Reference to tables count.</param>
    /// <param name="charts">Reference to charts count.</param>
    /// <param name="smartArt">Reference to SmartArt count.</param>
    /// <param name="audio">Reference to audio count.</param>
    /// <param name="video">Reference to video count.</param>
    public static void CountShapeTypes(IShape shape, ref int images, ref int tables, ref int charts,
        ref int smartArt, ref int audio, ref int video)
    {
        switch (shape)
        {
            case IAudioFrame:
                audio++;
                break;
            case IVideoFrame:
                video++;
                break;
            case PictureFrame:
                images++;
                break;
            case ITable:
                tables++;
                break;
            case IChart:
                charts++;
                break;
            case ISmartArt:
                smartArt++;
                break;
            case IGroupShape groupShape:
                foreach (var childShape in groupShape.Shapes)
                    CountShapeTypes(childShape, ref images, ref tables, ref charts, ref smartArt, ref audio, ref video);
                break;
        }
    }

    /// <summary>
    ///     Generates a thumbnail image of a slide as Base64 encoded PNG.
    /// </summary>
    /// <param name="slide">The slide to generate thumbnail from.</param>
    /// <param name="scaleX">Horizontal scale factor (default 0.5 = 50%).</param>
    /// <param name="scaleY">Vertical scale factor (default 0.5 = 50%).</param>
    /// <returns>Base64 encoded PNG image string.</returns>
    public static string GenerateThumbnail(ISlide slide, float scaleX = 0.5f, float scaleY = 0.5f)
    {
        using var bitmap = slide.GetThumbnail(scaleX, scaleY);
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return Convert.ToBase64String(stream.ToArray());
    }
}
