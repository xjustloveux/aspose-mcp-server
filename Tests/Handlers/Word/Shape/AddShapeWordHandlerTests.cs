using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Handlers.Word.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Shape;

public class AddShapeWordHandlerTests : WordHandlerTestBase
{
    private readonly AddShapeWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static List<Aspose.Words.Drawing.Shape> GetAllShapes(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true)
            .Cast<Aspose.Words.Drawing.Shape>()
            .ToList();
    }

    #endregion

    #region Basic Add Shape Operations

    [Fact]
    public void Execute_AddsRectangleShape()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" },
            { "width", 100.0 },
            { "height", 50.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var shapes = GetAllShapes(doc);
        Assert.NotEmpty(shapes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains(shapes, s => s.ShapeType == ShapeType.Rectangle);
            var rect = shapes.First(s => s.ShapeType == ShapeType.Rectangle);
            Assert.Equal(100.0, rect.Width);
            Assert.Equal(50.0, rect.Height);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsOvalShape()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Ellipse" },
            { "width", 80.0 },
            { "height", 60.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var shapes = GetAllShapes(doc);
        Assert.NotEmpty(shapes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains(shapes, s => s.ShapeType == ShapeType.Ellipse);
            var ellipse = shapes.First(s => s.ShapeType == ShapeType.Ellipse);
            Assert.Equal(80.0, ellipse.Width);
            Assert.Equal(60.0, ellipse.Height);
        }
    }

    [Fact]
    public void Execute_WithPosition_AddsShapeAtPosition()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" },
            { "width", 100.0 },
            { "height", 50.0 },
            { "x", 150.0 },
            { "y", 200.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var shapes = GetAllShapes(doc);
        Assert.NotEmpty(shapes);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var shape = shapes.First(s => s.ShapeType == ShapeType.Rectangle);
            Assert.Equal(150.0, shape.Left);
            Assert.Equal(200.0, shape.Top);
        }
    }

    [Fact]
    public void Execute_WithoutShapeType_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 100.0 },
            { "height", 50.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutWidth_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" },
            { "height", 50.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutHeight_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeType", "Rectangle" },
            { "width", 100.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
