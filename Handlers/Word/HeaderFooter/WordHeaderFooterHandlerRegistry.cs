using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public static class WordHeaderFooterHandlerRegistry
{
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new SetHeaderTextHandler());
        registry.Register(new SetFooterTextHandler());
        registry.Register(new SetHeaderImageHandler());
        registry.Register(new SetFooterImageHandler());
        registry.Register(new SetHeaderLineHandler());
        registry.Register(new SetFooterLineHandler());
        registry.Register(new SetHeaderTabsHandler());
        registry.Register(new SetFooterTabsHandler());
        registry.Register(new SetHeaderFooterHandler());
        registry.Register(new GetHeadersFootersHandler());
        return registry;
    }
}
