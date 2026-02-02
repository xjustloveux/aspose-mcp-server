using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IConnector connector)
            return null;

        string? lineColor = null;
        double? lineWidth = null;
        string? lineDashStyle = null;
        var lineFormat = connector.LineFormat;
        if (lineFormat != null)
        {
            if (lineFormat.FillFormat is { FillType: FillType.Solid })
            {
                var lc = lineFormat.FillFormat.SolidFillColor.Color;
                if (lc != Color.Empty)
                    lineColor = $"#{lc.R:X2}{lc.G:X2}{lc.B:X2}";
            }

            if (lineFormat.Width is > 0 and not double.NaN)
                lineWidth = lineFormat.Width;

            lineDashStyle = lineFormat.DashStyle.ToString();
            if (lineDashStyle == "NotDefined")
                lineDashStyle = null;
        }

        return new ConnectorDetails
        {
            ConnectorType = connector.ShapeType.ToString(),
            StartShapeConnectedTo = connector.StartShapeConnectedTo?.Name,
            EndShapeConnectedTo = connector.EndShapeConnectedTo?.Name,
            StartShapeConnectionSiteIndex = connector.StartShapeConnectionSiteIndex,
            EndShapeConnectionSiteIndex = connector.EndShapeConnectionSiteIndex,
            LineColor = lineColor,
            LineWidth = lineWidth,
            LineDashStyle = lineDashStyle
        };
    }
}
