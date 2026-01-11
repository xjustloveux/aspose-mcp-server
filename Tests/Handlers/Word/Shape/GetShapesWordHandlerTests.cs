using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class GetShapesWordHandlerTests : WordHandlerTestBase
{
    private readonly GetShapesWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
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

    #region Basic Get Shapes Operations

    [Fact]
    public void Execute_WithNoShapes_ReturnsNoShapesMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no shapes found", result.ToLower());
    }

    [Fact]
    public void Execute_WithShapes_ReturnsShapeInfo()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("total shapes:", result.ToLower());
        Assert.Contains("type:", result.ToLower());
    }

    [Fact]
    public void Execute_ReturnsSizeAndPosition()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("size:", result.ToLower());
        Assert.Contains("position:", result.ToLower());
    }

    #endregion
}
