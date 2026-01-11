using AsposeMcpServer.Handlers.Pdf.Attachment;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Attachment;

public class PdfAttachmentHandlerRegistryTests
{
    [Fact]
    public void Create_ReturnsValidRegistry()
    {
        var registry = PdfAttachmentHandlerRegistry.Create();

        Assert.NotNull(registry);
    }

    [Theory]
    [InlineData("add")]
    [InlineData("delete")]
    [InlineData("get")]
    public void GetHandler_WithValidOperation_ReturnsHandler(string operation)
    {
        var registry = PdfAttachmentHandlerRegistry.Create();

        var handler = registry.GetHandler(operation);

        Assert.NotNull(handler);
        Assert.Equal(operation, handler.Operation);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("GET")]
    [InlineData("Get")]
    public void GetHandler_IsCaseInsensitive(string operation)
    {
        var registry = PdfAttachmentHandlerRegistry.Create();

        var handler = registry.GetHandler(operation);

        Assert.NotNull(handler);
    }

    [Fact]
    public void GetHandler_WithInvalidOperation_ThrowsArgumentException()
    {
        var registry = PdfAttachmentHandlerRegistry.Create();

        Assert.Throws<ArgumentException>(() => registry.GetHandler("invalid"));
    }

    [Fact]
    public void GetHandler_WithEmptyOperation_ThrowsArgumentException()
    {
        var registry = PdfAttachmentHandlerRegistry.Create();

        Assert.Throws<ArgumentException>(() => registry.GetHandler(""));
    }
}
