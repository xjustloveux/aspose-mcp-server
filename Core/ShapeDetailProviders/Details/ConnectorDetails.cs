using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for Connector elements.
/// </summary>
public sealed record ConnectorDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the connector type.
    /// </summary>
    [JsonPropertyName("connectorType")]
    public required string ConnectorType { get; init; }

    /// <summary>
    ///     Gets the name of the shape connected at the start.
    /// </summary>
    [JsonPropertyName("startShapeConnectedTo")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StartShapeConnectedTo { get; init; }

    /// <summary>
    ///     Gets the name of the shape connected at the end.
    /// </summary>
    [JsonPropertyName("endShapeConnectedTo")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? EndShapeConnectedTo { get; init; }

    /// <summary>
    ///     Gets the connection site index at the start shape.
    /// </summary>
    [JsonPropertyName("startShapeConnectionSiteIndex")]
    public required uint StartShapeConnectionSiteIndex { get; init; }

    /// <summary>
    ///     Gets the connection site index at the end shape.
    /// </summary>
    [JsonPropertyName("endShapeConnectionSiteIndex")]
    public required uint EndShapeConnectionSiteIndex { get; init; }

    /// <summary>
    ///     Gets the line color in #RRGGBB format.
    /// </summary>
    [JsonPropertyName("lineColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LineColor { get; init; }

    /// <summary>
    ///     Gets the line width in points.
    /// </summary>
    [JsonPropertyName("lineWidth")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? LineWidth { get; init; }

    /// <summary>
    ///     Gets the line dash style (e.g., Solid, Dash, Dot).
    /// </summary>
    [JsonPropertyName("lineDashStyle")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LineDashStyle { get; init; }
}
