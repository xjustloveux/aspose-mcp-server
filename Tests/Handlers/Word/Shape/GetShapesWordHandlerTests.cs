using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Results.Word.Shape;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesWordResult>(res);

        Assert.Contains("no shapes found", result.Content.ToLower());
    }

    [Fact]
    public void Execute_WithShapes_ReturnsShapeInfo()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesWordResult>(res);

        Assert.Contains("total shapes:", result.Content.ToLower());
        Assert.Contains("type:", result.Content.ToLower());
    }

    [Fact]
    public void Execute_ReturnsSizeAndPosition()
    {
        var doc = CreateDocumentWithShape();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesWordResult>(res);

        Assert.Contains("size:", result.Content.ToLower());
        Assert.Contains("position:", result.Content.ToLower());
    }

    #endregion
}
