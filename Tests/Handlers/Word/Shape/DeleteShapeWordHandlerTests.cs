using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class DeleteShapeWordHandlerTests : WordHandlerTestBase
{
    private readonly DeleteShapeWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithShape()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        return doc;
    }

    #endregion

    #region Basic Delete Shape Operations

    [Fact]
    public void Execute_DeletesShape()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("successfully deleted", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
