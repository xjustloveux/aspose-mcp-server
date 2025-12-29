using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptNotesToolTests : TestBase
{
    private readonly PptNotesTool _tool = new();

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithNotes(string fileName, string notesText)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var notesSlide = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
        if (notesSlide.NotesTextFrame != null)
            notesSlide.NotesTextFrame.Text = notesText;
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region Unknown Operation Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Set Tests

    [Fact]
    public async Task Set_ShouldSetNotes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_notes.pptx");
        var outputPath = CreateTestFilePath("test_set_notes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["notes"] = "Speaker notes for this slide"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Notes set", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: content may be truncated with watermark
            Assert.False(string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text));
        else
            // Licensed mode: strict content verification
            Assert.Contains("Speaker notes", notesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public async Task Set_ShouldReplaceExistingNotes()
    {
        // Arrange
        var pptPath = CreatePresentationWithNotes("test_set_replace.pptx", "Original notes");
        var outputPath = CreateTestFilePath("test_set_replace_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["notes"] = "Updated notes content"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Notes set", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: just verify notes were changed (not equal to original)
            Assert.DoesNotContain("Original notes", notesSlide.NotesTextFrame.Text);
        else
            // Licensed mode: strict content verification
            Assert.Equal("Updated notes content", notesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public async Task Set_WithInvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["slideIndex"] = 99,
            ["notes"] = "Test notes"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    #endregion

    #region Get Tests

    [Fact]
    public async Task Get_WithSlideIndex_ShouldReturnSingleSlideNotes()
    {
        // Arrange
        var pptPath = CreatePresentationWithNotes("test_get_notes.pptx", "Test notes content");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: content may be truncated with watermark
            Assert.False(string.IsNullOrEmpty(json.RootElement.GetProperty("notes").GetString()));
        else
            // Licensed mode: strict content verification
            Assert.Contains("Test notes", json.RootElement.GetProperty("notes").GetString());
    }

    [Fact]
    public async Task Get_WithoutSlideIndex_ShouldReturnAllNotes()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_all_notes.pptx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("slides").GetArrayLength());
    }

    [Fact]
    public async Task Get_WithInvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 99
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    #endregion

    #region Clear Tests

    [Fact]
    public async Task Clear_WithSlideIndices_ShouldClearSpecificSlides()
    {
        // Arrange
        var pptPath = CreatePresentationWithNotes("test_clear_notes.pptx", "Notes to clear");
        var outputPath = CreateTestFilePath("test_clear_notes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "clear",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndices"] = new JsonArray { 0 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Cleared", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.True(string.IsNullOrEmpty(notesSlide?.NotesTextFrame?.Text));
    }

    [Fact]
    public async Task Clear_WithoutSlideIndices_ShouldClearAllSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_clear_all.pptx", 3);
        var outputPath = CreateTestFilePath("test_clear_all_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "clear",
            ["path"] = pptPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("3 targeted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task Clear_WithInvalidSlideIndices_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_clear_invalid.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "clear",
            ["path"] = pptPath,
            ["slideIndices"] = new JsonArray { 0, 99 }
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    #endregion

    #region SetHeaderFooter Tests

    [Fact]
    public async Task SetHeaderFooter_ShouldSetNotesHeaderFooter()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_notes_header_footer.pptx");
        var outputPath = CreateTestFilePath("test_notes_header_footer_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["headerText"] = "Notes Header",
            ["footerText"] = "Notes Footer"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetHeaderFooter_WithAllParameters_ShouldSetAll()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_notes_hf_all.pptx");
        var outputPath = CreateTestFilePath("test_notes_hf_all_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["headerText"] = "Header",
            ["footerText"] = "Footer",
            ["dateText"] = "2024-12-28",
            ["showPageNumber"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        Assert.Contains("date", result);
        Assert.Contains("page number shown", result);
    }

    [Fact]
    public async Task SetHeaderFooter_HidePageNumber_ShouldHide()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_notes_hide_page.pptx");
        var outputPath = CreateTestFilePath("test_notes_hide_page_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["showPageNumber"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("page number hidden", result);
    }

    [Fact]
    public async Task SetHeaderFooter_OnEmptyPresentation_ShouldThrow()
    {
        // Arrange - Create empty presentation
        var pptPath = CreateTestFilePath("test_notes_empty.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.RemoveAt(0);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["headerText"] = "Header"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("no slides", ex.Message);
    }

    #endregion
}