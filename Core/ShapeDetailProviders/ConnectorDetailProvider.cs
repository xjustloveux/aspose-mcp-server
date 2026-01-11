using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from Connector elements
/// </summary>
public class ConnectorDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Connector";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IConnector;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IConnector connector)
            return null;

        return new
        {
            connectorType = connector.ShapeType.ToString(),
            startShapeConnectedTo = connector.StartShapeConnectedTo?.Name,
            endShapeConnectedTo = connector.EndShapeConnectedTo?.Name,
            startShapeConnectionSiteIndex = connector.StartShapeConnectionSiteIndex,
            endShapeConnectionSiteIndex = connector.EndShapeConnectionSiteIndex
        };
    }
}
