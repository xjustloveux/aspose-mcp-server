using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using Shape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>
///     Diagnostic asserting which of the three "linked" fixtures actually produces an
///     <c>IsLink=true</c> OLE on reload. Used by the test-engineer to scope the linked
///     AC assertions correctly (see the <see cref="FixtureKind" /> enumeration). Not an
///     AC test itself — a contract check consumed by <c>OleErrorParityTests</c> and
///     <c>*OleObjectToolTests</c>.
/// </summary>
[Collection(OleFixtureCollection.Name)]
public sealed class FixtureIsLinkDiagTests
{
    private readonly FixtureBuilder _fixtures;

    /// <summary>Initializes a new instance of the <see cref="FixtureIsLinkDiagTests" /> class.</summary>
    /// <param name="fixtures">Shared fixture matrix.</param>
    public FixtureIsLinkDiagTests(FixtureBuilder fixtures)
    {
        _fixtures = fixtures;
    }

    /// <summary>Confirms the Word linked fixture produces an IsLink=true OLE on reload.</summary>
    [Fact]
    public void WordLinkedDocx_IsLinkTrue()
    {
        var doc = new Document(_fixtures.Paths[FixtureKind.WordLinkedDocx]);
        var shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        Assert.True(shape.OleFormat!.IsLink);
    }

    /// <summary>Reports the Excel linked fixture's IsLink state — informational assertion only.</summary>
    [Fact]
    public void ExcelLinkedXlsx_IsLinkStateIsKnown()
    {
        using var wb = new Workbook(_fixtures.Paths[FixtureKind.ExcelLinkedXlsx]);
        var ole = wb.Worksheets[0].OleObjects[0];
        // Aspose.Cells 23.10.0 treats OLE-with-payload as embedded even when
        // ObjectSourceFullName is set, so IsLink returns false on reload.
        // Tests that require a reliably linked Excel fixture should use a different vector.
        // This assertion pins the known observable state so downstream tests can depend on it.
        Assert.False(ole.IsLink,
            "Aspose.Cells 23.10.0 reports IsLink=false for the ExcelLinkedXlsx fixture " +
            "because OLE-with-payload is treated as embedded; see class summary.");
    }

    /// <summary>Confirms the PPT linked fixture produces an IsObjectLink=true frame on reload.</summary>
    [Fact]
    public void PptLinkedPptx_IsObjectLinkTrue()
    {
        using var pres = new Presentation(_fixtures.Paths[FixtureKind.PptLinkedPptx]);
        var frame = (IOleObjectFrame)pres.Slides[0].Shapes[0];
        Assert.True(frame.IsObjectLink);
    }
}
