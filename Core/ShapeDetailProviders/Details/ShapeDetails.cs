using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Abstract base type for shape-specific detail information.
///     Uses JSON polymorphism to serialize/deserialize derived types with a type discriminator.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "$type")]
[JsonDerivedType(typeof(AutoShapeDetails), "autoShape")]
[JsonDerivedType(typeof(AudioFrameDetails), "audio")]
[JsonDerivedType(typeof(ChartDetails), "chart")]
[JsonDerivedType(typeof(ConnectorDetails), "connector")]
[JsonDerivedType(typeof(GroupShapeDetails), "group")]
[JsonDerivedType(typeof(PictureFrameDetails), "picture")]
[JsonDerivedType(typeof(SmartArtDetails), "smartArt")]
[JsonDerivedType(typeof(TableDetails), "table")]
[JsonDerivedType(typeof(VideoFrameDetails), "video")]
public abstract record ShapeDetails;
