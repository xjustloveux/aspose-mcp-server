using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptNotesToolTests : TestBase
{
    private readonly PptNotesTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddNotes_ShouldAddNotes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_notes.pptx");
        var outputPath = CreateTestFilePath("test_add_notes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["notes"] = "Speaker notes for this slide"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var notesSlide = slide.NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
    }

    [Fact]
    public async Task GetNotes_ShouldReturnNotes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_notes.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
            if (notesSlide.NotesTextFrame != null) notesSlide.NotesTextFrame.Text = "Test notes";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Notes", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ClearNotes_ShouldClearNotes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_clear_notes.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
            if (notesSlide.NotesTextFrame != null) notesSlide.NotesTextFrame.Text = "Notes to clear";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_clear_notes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "clear",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndices"] = new JsonArray { 0 }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }
}