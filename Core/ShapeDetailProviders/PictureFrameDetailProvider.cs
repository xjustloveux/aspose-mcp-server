using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from PictureFrame elements
/// </summary>
public class PictureFrameDetailProvider : IShapeDetailProvider
{
    public string TypeName => "Picture";

    public bool CanHandle(IShape shape)
    {
        return shape is IPictureFrame;
    }

    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IPictureFrame picture)
            return null;

        var image = picture.PictureFormat?.Picture?.Image;

        return new
        {
            alternativeText = picture.AlternativeText,
            contentType = image?.ContentType,
            imageWidth = image?.Width,
            imageHeight = image?.Height,
            cropLeft = picture.PictureFormat?.CropLeft,
            cropRight = picture.PictureFormat?.CropRight,
            cropTop = picture.PictureFormat?.CropTop,
            cropBottom = picture.PictureFormat?.CropBottom
        };
    }
}