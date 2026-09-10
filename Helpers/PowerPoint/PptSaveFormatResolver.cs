using Aspose.Slides.Export;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Chooses the presentation save format from the destination file's extension.
///     Every save path previously passed <see cref="SaveFormat.Pptx" /> unconditionally, so editing
///     a <c>.ppt</c>, <c>.pptm</c>, <c>.potx</c> or <c>.odp</c> wrote PPTX content into a file whose
///     name still claimed the original format. A macro-enabled presentation lost its macros that
///     way, and the caller was told the save succeeded. Word, Excel and PDF already infer the
///     format from the path; this brings PowerPoint in line.
/// </summary>
public static class PptSaveFormatResolver
{
    /// <summary>
    ///     Extensions this library can write, mapped to the format that preserves them.
    /// </summary>
    private static readonly Dictionary<string, SaveFormat> FormatsByExtension =
        new(StringComparer.OrdinalIgnoreCase)
        {
            [".pptx"] = SaveFormat.Pptx,
            [".ppt"] = SaveFormat.Ppt,
            [".pptm"] = SaveFormat.Pptm,
            [".potx"] = SaveFormat.Potx,
            [".potm"] = SaveFormat.Potm,
            [".pot"] = SaveFormat.Pot,
            [".odp"] = SaveFormat.Odp,
            [".otp"] = SaveFormat.Otp,
            [".ppsx"] = SaveFormat.Ppsx,
            [".ppsm"] = SaveFormat.Ppsm,
            [".pps"] = SaveFormat.Pps,
            [".xps"] = SaveFormat.Xps,
            [".pdf"] = SaveFormat.Pdf,
            [".html"] = SaveFormat.Html,
            [".htm"] = SaveFormat.Html,
            [".tiff"] = SaveFormat.Tiff,
            [".gif"] = SaveFormat.Gif
        };

    /// <summary>
    ///     Resolves the save format for a destination path.
    /// </summary>
    /// <param name="path">Destination file path; only its extension is inspected.</param>
    /// <returns>The format that writes content matching the extension.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the extension is missing or is one this library cannot write. Failing here is
    ///     deliberate: writing a different format under that name would corrupt the file silently.
    /// </exception>
    public static SaveFormat Resolve(string path)
    {
        var extension = Path.GetExtension(path);
        if (string.IsNullOrEmpty(extension))
            throw new ArgumentException(
                $"Cannot determine the presentation format for '{Path.GetFileName(path)}' because it has no extension.");

        if (FormatsByExtension.TryGetValue(extension, out var format))
            return format;

        throw new ArgumentException(
            $"Presentations cannot be saved as '{extension}'. Supported extensions: "
            + string.Join(", ", FormatsByExtension.Keys.Order(StringComparer.OrdinalIgnoreCase)) + ".");
    }
}
