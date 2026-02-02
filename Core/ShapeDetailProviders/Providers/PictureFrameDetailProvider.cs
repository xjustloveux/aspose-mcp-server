using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

/// <summary>
///     Provider for extracting details from PictureFrame elements
/// </summary>
public class PictureFrameDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Picture";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IPictureFrame;
    }

    /// <inheritdoc />
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IPictureFrame picture)
            return null;

        var image = picture.PictureFormat?.Picture?.Image;

        return new PictureFrameDetails
        {
            Hyperlink = AutoShapeDetailProvider.GetHyperlinkText(picture.HyperlinkClick, presentation),
            ContentType = image?.ContentType,
            ImageWidth = image?.Width,
            ImageHeight = image?.Height,
            CropLeft = picture.PictureFormat?.CropLeft,
            CropRight = picture.PictureFormat?.CropRight,
            CropTop = picture.PictureFormat?.CropTop,
            CropBottom = picture.PictureFormat?.CropBottom
        };
    }
}
