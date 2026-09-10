using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Tests.Handlers.Word.Image;

public class DeleteImageWordHandlerTests : WordHandlerTestBase
{
    private readonly DeleteImageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesImage()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var initialCount = GetImageCount(doc);
        Assert.Equal(1, initialCount);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, GetImageCount(doc));
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsImageIndex()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithImage(tempFile);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#0", result.Message);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateDocumentWithMultipleImages(tempFile, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Remaining images: 2", result.Message);
    }

    /// <summary>
    ///     Three identical images could not show which one a delete removed, so an off-by-one in
    ///     the index would still have passed. Distinct images make the survivors identifiable.
    /// </summary>
    [Fact]
    public void Execute_ShouldDeleteTheAddressedImageAndKeepTheOrderOfTheRest()
    {
        var first = CreateTempImageFile(10);
        var second = CreateTempImageFile(120);
        var third = CreateTempImageFile(240);
        var doc = CreateDocumentWithDistinctImages(first, second, third);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?> { { "imageIndex", 1 } });

        var sizesBefore = ImageByteLengths(doc);
        _handler.Execute(context, parameters);
        var sizesAfter = ImageByteLengths(doc);

        Assert.Equal(3, sizesBefore.Count);
        Assert.Equal(2, sizesAfter.Count);
        Assert.Equal([sizesBefore[0], sizesBefore[2]], sizesAfter);
    }

    /// <summary>
    ///     Reads each image's stored bytes so images can be compared by content.
    /// </summary>
    /// <param name="doc">The document to read.</param>
    /// <returns>The image bytes, in document order.</returns>
    private static List<string> ImageByteLengths(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .Where(shape => shape.HasImage)
            .Select(shape => Convert.ToBase64String(shape.ImageData.ImageBytes))
            .ToList();
    }

    #endregion

    #region Error Handling

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    [InlineData(10)]
    public void Execute_WithInvalidImageIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyDocument_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imageIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithImage(string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        return doc;
    }

    private static Document CreateDocumentWithMultipleImages(string imagePath, int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
        }

        return doc;
    }

    /// <summary>
    ///     Builds a document holding one image per supplied path, so the caller can identify which
    ///     image survived a deletion.
    /// </summary>
    /// <param name="imagePaths">Distinct image files, in insertion order.</param>
    /// <returns>The document.</returns>
    private static Document CreateDocumentWithDistinctImages(params string[] imagePaths)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        foreach (var path in imagePaths)
        {
            builder.InsertImage(path);
            builder.InsertParagraph();
        }

        return doc;
    }

    private static int GetImageCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .Count(s => s.HasImage);
    }

    #endregion
}
