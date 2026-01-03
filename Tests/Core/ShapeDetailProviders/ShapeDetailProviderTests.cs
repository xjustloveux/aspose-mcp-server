using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

/// <summary>
///     Unit tests for ShapeDetailProviderFactory and individual shape detail providers
/// </summary>
public class ShapeDetailProviderFactoryTests : TestBase
{
    #region GetProvider Tests

    [Fact]
    public void GetProvider_WithAutoShape_ShouldReturnAutoShapeProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var provider = ShapeDetailProviderFactory.GetProvider(shape);

        Assert.NotNull(provider);
        Assert.IsType<AutoShapeDetailProvider>(provider);
    }

    [Fact]
    public void GetProvider_WithTable_ShouldReturnTableProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var provider = ShapeDetailProviderFactory.GetProvider(table);

        Assert.NotNull(provider);
        Assert.IsType<TableDetailProvider>(provider);
    }

    [Fact]
    public void GetProvider_WithConnector_ShouldReturnConnectorProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var provider = ShapeDetailProviderFactory.GetProvider(connector);

        Assert.NotNull(provider);
        Assert.IsType<ConnectorDetailProvider>(provider);
    }

    #endregion

    #region GetShapeDetails Tests

    [Fact]
    public void GetShapeDetails_WithAutoShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Test Text";

        var (typeName, properties) = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

        Assert.Equal("AutoShape", typeName);
        Assert.NotNull(properties);
    }

    [Fact]
    public void GetShapeDetails_WithTable_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var (typeName, properties) = ShapeDetailProviderFactory.GetShapeDetails(table, presentation);

        Assert.Equal("Table", typeName);
        Assert.NotNull(properties);
    }

    #endregion
}

/// <summary>
///     Unit tests for AutoShapeDetailProvider
/// </summary>
public class AutoShapeDetailProviderTests : TestBase
{
    private readonly AutoShapeDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnAutoShape()
    {
        Assert.Equal("AutoShape", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithTable_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var result = _provider.CanHandle(table);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithAutoShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Sample Text";

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonAutoShape_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var details = _provider.GetDetails(table, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithEmptyTextFrame_ShouldIncludeNullText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }
}

/// <summary>
///     Unit tests for TableDetailProvider
/// </summary>
public class TableDetailProviderTests : TestBase
{
    private readonly TableDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnTable()
    {
        Assert.Equal("Table", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithTable_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var result = _provider.CanHandle(table);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithTable_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonTable_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }
}

/// <summary>
///     Unit tests for ConnectorDetailProvider
/// </summary>
public class ConnectorDetailProviderTests : TestBase
{
    private readonly ConnectorDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnConnector()
    {
        Assert.Equal("Connector", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithConnector_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var result = _provider.CanHandle(connector);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithConnector_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var details = _provider.GetDetails(connector, presentation);

        Assert.NotNull(details);
    }
}

/// <summary>
///     Unit tests for GroupShapeDetailProvider
/// </summary>
public class GroupShapeDetailProviderTests : TestBase
{
    private readonly GroupShapeDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnGroup()
    {
        Assert.Equal("Group", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void CanHandle_WithGroupShape_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var group = slide.Shapes.AddGroupShape();

        var result = _provider.CanHandle(group);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithGroupShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var group = slide.Shapes.AddGroupShape();

        var details = _provider.GetDetails(group, presentation);

        Assert.NotNull(details);
    }
}

/// <summary>
///     Unit tests for PictureFrameDetailProvider
/// </summary>
public class PictureFrameDetailProviderTests : TestBase
{
    private readonly PictureFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnPicture()
    {
        Assert.Equal("Picture", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }
}

/// <summary>
///     Unit tests for AudioFrameDetailProvider
/// </summary>
public class AudioFrameDetailProviderTests : TestBase
{
    private readonly AudioFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnAudio()
    {
        Assert.Equal("Audio", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }
}

/// <summary>
///     Unit tests for VideoFrameDetailProvider
/// </summary>
public class VideoFrameDetailProviderTests : TestBase
{
    private readonly VideoFrameDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnVideo()
    {
        Assert.Equal("Video", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }
}

/// <summary>
///     Unit tests for ChartDetailProvider
/// </summary>
public class ChartDetailProviderTests : TestBase
{
    private readonly ChartDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnChart()
    {
        Assert.Equal("Chart", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }
}

/// <summary>
///     Unit tests for SmartArtDetailProvider
/// </summary>
public class SmartArtDetailProviderTests : TestBase
{
    private readonly SmartArtDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnSmartArt()
    {
        Assert.Equal("SmartArt", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }
}