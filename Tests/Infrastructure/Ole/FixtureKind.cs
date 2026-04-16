namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>Enumeration of the 12 fixture kinds produced by <see cref="FixtureBuilder" />.</summary>
public enum FixtureKind
{
    /// <summary>Word .docx with one embedded OLE object.</summary>
    WordEmbeddedDocx,

    /// <summary>Word .docx with one linked OLE object.</summary>
    WordLinkedDocx,

    /// <summary>Word .docx carrying an attacker-controlled raw filename.</summary>
    WordAttackerDocx,

    /// <summary>Legacy Word .doc with one embedded OLE object.</summary>
    WordEmbeddedDoc,

    /// <summary>Excel .xlsx with one embedded OLE object.</summary>
    ExcelEmbeddedXlsx,

    /// <summary>Excel .xlsx with one linked OLE object.</summary>
    ExcelLinkedXlsx,

    /// <summary>Excel .xlsx carrying an attacker-controlled label.</summary>
    ExcelAttackerXlsx,

    /// <summary>Legacy Excel .xls with one embedded OLE object.</summary>
    ExcelEmbeddedXls,

    /// <summary>PowerPoint .pptx with one embedded OLE frame.</summary>
    PptEmbeddedPptx,

    /// <summary>PowerPoint .pptx with one linked OLE frame.</summary>
    PptLinkedPptx,

    /// <summary>PowerPoint .pptx carrying an attacker-controlled embedded filename.</summary>
    PptAttackerPptx,

    /// <summary>Legacy PowerPoint .ppt with one embedded OLE frame.</summary>
    PptEmbeddedPpt
}
