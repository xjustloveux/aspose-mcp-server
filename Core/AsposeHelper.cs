using Aspose.Cells;
using Aspose.Slides;
using AsposePdf = Aspose.Pdf;

namespace AsposeMcpServer.Core;

/// <summary>
/// Helper class for common Excel operations to reduce code duplication
/// </summary>
public static class ExcelHelper
{
    /// <summary>
    /// Validates sheet index and throws exception if invalid
    /// </summary>
    /// <param name="sheetIndex">Sheet index to validate</param>
    /// <param name="workbook">Workbook to check against</param>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static void ValidateSheetIndex(int sheetIndex, Workbook workbook)
    {
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }
    }

    /// <summary>
    /// Validates sheet index and throws exception if invalid (with custom error message)
    /// </summary>
    /// <param name="sheetIndex">Sheet index to validate</param>
    /// <param name="workbook">Workbook to check against</param>
    /// <param name="customMessage">Custom error message prefix</param>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static void ValidateSheetIndex(int sheetIndex, Workbook workbook, string customMessage)
    {
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"{customMessage}: 工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }
    }

    /// <summary>
    /// Gets a worksheet with validation
    /// </summary>
    /// <param name="workbook">Workbook to get worksheet from</param>
    /// <param name="sheetIndex">Sheet index</param>
    /// <returns>Worksheet</returns>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static Worksheet GetWorksheet(Workbook workbook, int sheetIndex)
    {
        ValidateSheetIndex(sheetIndex, workbook);
        return workbook.Worksheets[sheetIndex];
    }

    /// <summary>
    /// Gets a worksheet with validation (with custom error message)
    /// </summary>
    /// <param name="workbook">Workbook to get worksheet from</param>
    /// <param name="sheetIndex">Sheet index</param>
    /// <param name="customMessage">Custom error message prefix</param>
    /// <returns>Worksheet</returns>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static Worksheet GetWorksheet(Workbook workbook, int sheetIndex, string customMessage)
    {
        ValidateSheetIndex(sheetIndex, workbook, customMessage);
        return workbook.Worksheets[sheetIndex];
    }
}

/// <summary>
/// Helper class for common PowerPoint operations to reduce code duplication
/// </summary>
public static class PowerPointHelper
{
    /// <summary>
    /// Validates slide index and throws exception if invalid
    /// </summary>
    /// <param name="slideIndex">Slide index to validate</param>
    /// <param name="presentation">Presentation to check against</param>
    /// <exception cref="ArgumentException">Thrown if slide index is invalid</exception>
    public static void ValidateSlideIndex(int slideIndex, IPresentation presentation)
    {
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"投影片索引 {slideIndex} 超出範圍 (共有 {presentation.Slides.Count} 個投影片)");
        }
    }

    /// <summary>
    /// Gets a slide with validation
    /// </summary>
    /// <param name="presentation">Presentation to get slide from</param>
    /// <param name="slideIndex">Slide index</param>
    /// <returns>Slide</returns>
    /// <exception cref="ArgumentException">Thrown if slide index is invalid</exception>
    public static ISlide GetSlide(IPresentation presentation, int slideIndex)
    {
        ValidateSlideIndex(slideIndex, presentation);
        return presentation.Slides[slideIndex];
    }

    /// <summary>
    /// Validates shape index and throws exception if invalid
    /// </summary>
    /// <param name="shapeIndex">Shape index to validate</param>
    /// <param name="slide">Slide to check against</param>
    /// <exception cref="ArgumentException">Thrown if shape index is invalid</exception>
    public static void ValidateShapeIndex(int shapeIndex, ISlide slide)
    {
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"形狀索引 {shapeIndex} 超出範圍 (共有 {slide.Shapes.Count} 個形狀)");
        }
    }

    /// <summary>
    /// Gets a shape with validation
    /// </summary>
    /// <param name="slide">Slide to get shape from</param>
    /// <param name="shapeIndex">Shape index</param>
    /// <returns>Shape</returns>
    /// <exception cref="ArgumentException">Thrown if shape index is invalid</exception>
    public static IShape GetShape(ISlide slide, int shapeIndex)
    {
        ValidateShapeIndex(shapeIndex, slide);
        return slide.Shapes[shapeIndex];
    }

    /// <summary>
    /// Validates collection index and throws exception if invalid
    /// </summary>
    /// <typeparam name="T">Collection item type</typeparam>
    /// <param name="index">Index to validate</param>
    /// <param name="collection">Collection to check against</param>
    /// <param name="itemName">Name of the item type for error message</param>
    /// <exception cref="ArgumentException">Thrown if index is invalid</exception>
    public static void ValidateCollectionIndex<T>(int index, ICollection<T> collection, string itemName = "項目")
    {
        if (index < 0 || index >= collection.Count)
        {
            throw new ArgumentException($"{itemName}索引 {index} 超出範圍 (共有 {collection.Count} 個{itemName})");
        }
    }

    /// <summary>
    /// Validates collection index for collections with Count property (supports Aspose collections)
    /// </summary>
    /// <param name="index">Index to validate</param>
    /// <param name="count">Collection count</param>
    /// <param name="itemName">Name of the item type for error message</param>
    /// <exception cref="ArgumentException">Thrown if index is invalid</exception>
    public static void ValidateCollectionIndex(int index, int count, string itemName = "項目")
    {
        if (index < 0 || index >= count)
        {
            throw new ArgumentException($"{itemName}索引 {index} 超出範圍 (共有 {count} 個{itemName})");
        }
    }
}

/// <summary>
/// Helper class for common Word operations to reduce code duplication
/// </summary>
public static class WordHelper
{
    // Word document operations are typically simpler and don't require as much validation
    // Most operations work directly on the document without index validation
    // This class can be extended if common patterns emerge
}

/// <summary>
/// Helper class for common PDF operations to reduce code duplication
/// </summary>
public static class PdfHelper
{
    /// <summary>
    /// Validates page index and throws exception if invalid
    /// </summary>
    /// <param name="pageIndex">Page index to validate</param>
    /// <param name="document">Document to check against</param>
    /// <exception cref="ArgumentException">Thrown if page index is invalid</exception>
    public static void ValidatePageIndex(int pageIndex, AsposePdf.Document document)
    {
        if (pageIndex < 0 || pageIndex >= document.Pages.Count)
        {
            throw new ArgumentException($"頁面索引 {pageIndex} 超出範圍 (共有 {document.Pages.Count} 個頁面)");
        }
    }

    /// <summary>
    /// Gets a page with validation
    /// </summary>
    /// <param name="document">Document to get page from</param>
    /// <param name="pageIndex">Page index</param>
    /// <returns>Page</returns>
    /// <exception cref="ArgumentException">Thrown if page index is invalid</exception>
    public static AsposePdf.Page GetPage(AsposePdf.Document document, int pageIndex)
    {
        ValidatePageIndex(pageIndex, document);
        return document.Pages[pageIndex + 1]; // PDF pages are 1-based
    }

    /// <summary>
    /// Validates collection index and throws exception if invalid
    /// </summary>
    /// <typeparam name="T">Collection item type</typeparam>
    /// <param name="index">Index to validate</param>
    /// <param name="collection">Collection to check against</param>
    /// <param name="itemName">Name of the item type for error message</param>
    /// <exception cref="ArgumentException">Thrown if index is invalid</exception>
    public static void ValidateCollectionIndex<T>(int index, ICollection<T> collection, string itemName = "項目")
    {
        if (index < 0 || index >= collection.Count)
        {
            throw new ArgumentException($"{itemName}索引 {index} 超出範圍 (共有 {collection.Count} 個{itemName})");
        }
    }
}

