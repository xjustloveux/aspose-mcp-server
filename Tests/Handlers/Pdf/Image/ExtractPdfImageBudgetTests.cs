using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Tests.Infrastructure;
using Color = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

/// <summary>
///     R23-PDF01: extracting every image on a page is one request with one output budget — a
///     count admitted before the first write, bytes charged across all the images — and over
///     either budget nothing at all is published.
/// </summary>
[Collection("SerialStaticSeams")]
public class ExtractPdfImageBudgetTests : PdfHandlerTestBase
{
    private readonly ExtractPdfImageHandler _handler = new();

    private Document CreateDocumentWithImages(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var imagePath = Path.Combine(TestDir, $"budget_{Guid.NewGuid():N}.png");
        using (var bitmap = new Bitmap(120, 120))
        {
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Blue);
            }

            bitmap.Save(imagePath, ImageFormat.Png);
        }

        for (var i = 0; i < count; i++)
            page.AddImage(imagePath, new Rectangle(50 + i * 10, 500, 200 + i * 10, 650));
        return doc;
    }

    private string AnOutputDirectory()
    {
        var directory = Path.Combine(TestDir, "extracted_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    [Fact]
    public void MoreImagesThanTheFileBudget_AreRefusedBeforeTheFirstWrite()
    {
        var doc = CreateDocumentWithImages(3);
        var directory = AnOutputDirectory();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["pageIndex"] = 1,
            ["outputDir"] = directory
        });

        var before = ExtractPdfImageHandler.OutputFileBudget;
        ExtractPdfImageHandler.OutputFileBudget = 2;
        try
        {
            var refusal = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
            Assert.Contains("image files", refusal.Message, StringComparison.Ordinal);
        }
        finally
        {
            ExtractPdfImageHandler.OutputFileBudget = before;
        }

        Assert.Empty(Directory.GetFiles(directory));
    }

    [Fact]
    public void ImagesWhoseTotalExceedsTheByteBudget_AreRefused_WithNothingPublished()
    {
        // Each image alone is under the budget; together they are over it. The old code gave
        // each its own publisher with the full budget and published the first ones.
        var doc = CreateDocumentWithImages(3);
        var directory = AnOutputDirectory();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["pageIndex"] = 1,
            ["outputDir"] = directory
        });

        var before = ExtractPdfImageHandler.OutputByteBudget;
        // Measure one image by extracting all three under the real budget first, then set the
        // budget to just under the total of the three.
        _handler.Execute(context, parameters);
        var total = Directory.GetFiles(directory).Sum(f => new FileInfo(f).Length);
        foreach (var f in Directory.GetFiles(directory)) File.Delete(f);
        Assert.True(total > 3, "the images produced no bytes to budget");

        // A fresh document for the bounded run: the vendor's image objects do not survive a
        // second save, and that is not what is under test.
        var again = CreateContext(CreateDocumentWithImages(3));
        ExtractPdfImageHandler.OutputByteBudget = total - 1;
        try
        {
            Assert.Throws<ArgumentException>(() => _handler.Execute(again, parameters));
        }
        finally
        {
            ExtractPdfImageHandler.OutputByteBudget = before;
        }

        Assert.Empty(Directory.GetFiles(directory));
    }

    [Fact]
    public void ImagesWithinBothBudgets_AreAllPublished()
    {
        var doc = CreateDocumentWithImages(3);
        var directory = AnOutputDirectory();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["pageIndex"] = 1,
            ["outputDir"] = directory
        });

        _handler.Execute(context, parameters);

        Assert.Equal(3, Directory.GetFiles(directory, "*.png").Length);
    }
}
