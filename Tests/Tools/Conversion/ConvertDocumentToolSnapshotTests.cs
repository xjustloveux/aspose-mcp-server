using Aspose.Pdf;
using Aspose.Slides;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Conversion;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

/// <summary>
///     R23-PPT01: the presentation the preflight measured is the one the loader reads. The
///     caller's path is snapshotted once; swap the path afterwards and the conversion still
///     produces the admitted document.
/// </summary>
[Collection("SerialStaticSeams")]
public class ConvertDocumentToolSnapshotTests : TestBase
{
    private string APresentationWith(int slides, string name)
    {
        var path = CreateTestFilePath(name);
        using var presentation = new Presentation();
        for (var i = 1; i < slides; i++)
            presentation.Slides.AddClone(presentation.Slides[0]);
        presentation.Save(path, SlidesSaveFormat.Pptx);
        return path;
    }

    [Fact]
    public void APathSwappedAfterTheSnapshot_IsNotWhatTheLoaderReads()
    {
        var input = APresentationWith(1, "admitted.pptx");
        var replacement = APresentationWith(3, "replacement.pptx");
        var output = CreateTestFilePath("snapshot.pdf");
        var tool = new ConvertDocumentTool(SessionManager);

        // Fires once the snapshot exists and before anything reads it: the moment a caller
        // would swap the file between the preflight and the loader.
        ImmutableInputCopy.AfterCopy = _ => File.Copy(replacement, input, true);
        try
        {
            tool.Execute(input, outputPath: output);
        }
        finally
        {
            ImmutableInputCopy.AfterCopy = null;
        }

        using var pdf = new Document(output);
        Assert.Single(pdf.Pages);
    }
}
