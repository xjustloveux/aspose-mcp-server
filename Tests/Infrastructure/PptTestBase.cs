using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for PowerPoint tool tests providing PowerPoint-specific functionality.
/// </summary>
public abstract class PptTestBase : TestBase
{
    /// <summary>
    ///     Creates a new PowerPoint presentation with the specified number of slides.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <param name="slideCount">The number of slides to create (default: 1).</param>
    /// <returns>The file path of the created presentation.</returns>
    protected string CreatePresentation(string fileName, int slideCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with a single slide containing the specified text content.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <param name="content">The text content to add to the first slide.</param>
    /// <returns>The file path of the created presentation.</returns>
    protected string CreatePresentationWithContent(string fileName, string content)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 500, 100);
        shape.TextFrame.Text = content;
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with multiple slides, each containing specified text content.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <param name="slideContents">The text content for each slide.</param>
    /// <returns>The file path of the created presentation.</returns>
    protected string CreatePresentationWithSlides(string fileName, params string[] slideContents)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();

        for (var i = 0; i < slideContents.Length; i++)
        {
            var slide = i == 0
                ? presentation.Slides[0]
                : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 600, 100);
            shape.TextFrame.Text = slideContents[i];
        }

        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with a shape on the first slide.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <param name="slideCount">The number of slides to create (default: 1).</param>
    /// <returns>The file path of the created presentation.</returns>
    protected string CreatePresentationWithShape(string fileName, int slideCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a PowerPoint presentation with a table on the first slide.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <param name="rows">The number of rows in the table (default: 2).</param>
    /// <param name="columns">The number of columns in the table (default: 2).</param>
    /// <returns>The file path of the created presentation.</returns>
    protected string CreatePresentationWithTable(string fileName, int rows = 2, int columns = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var colWidths = Enumerable.Repeat(100.0, columns).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < columns; c++)
            table[c, r].TextFrame.Text = $"R{r}C{c}";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Asserts that a presentation has the expected number of slides.
    /// </summary>
    /// <param name="presentation">The presentation to check.</param>
    /// <param name="expectedCount">The expected number of slides.</param>
    protected static void AssertSlideCount(Presentation presentation, int expectedCount)
    {
        Assert.Equal(expectedCount, presentation.Slides.Count);
    }

    /// <summary>
    ///     Asserts that a slide has a shape at the specified index.
    /// </summary>
    /// <param name="slide">The slide to check.</param>
    /// <param name="shapeIndex">The expected shape index.</param>
    protected static void AssertSlideHasShape(ISlide slide, int shapeIndex)
    {
        Assert.True(shapeIndex < slide.Shapes.Count, $"Shape index {shapeIndex} is out of range");
    }

    /// <summary>
    ///     Checks if Aspose.Slides is running in evaluation mode.
    /// </summary>
    /// <param name="libraryType">The Aspose library type to check (default: Slides).</param>
    /// <returns>True if running in evaluation mode, false if licensed.</returns>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Slides)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}
