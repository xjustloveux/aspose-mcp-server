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

    private static int GetImageCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<WordShape>()
            .Count(s => s.HasImage);
    }

    #endregion
}
