using Microsoft.Office.Interop.PowerPoint;
using System.ComponentModel;

namespace LaRottaO.OfficeTranslationTool.Models
{
    internal class PptShape
    {
        [Browsable(false)]
        public int indexOnPresentation { get; set; }

        [Browsable(false)]
        public int slideNumber { get; set; }

        [Browsable(false)]
        public int indexOnSlide { get; set; }

        [Browsable(false)]
        public int internalId { get; set; }

        [Browsable(false)]
        public Boolean belongsToATable { get; set; }

        [Browsable(false)]
        public int parentTableRow { get; set; }

        [Browsable(false)]
        public int parentTableColumn { get; set; }

        [ColumnName("Info")]
        public String info { get; set; } = String.Empty;

        [ColumnName("Original Text")]
        public String originalText { get; set; } = String.Empty;

        [ColumnName("New Text")]
        public String newText { get; set; } = String.Empty;

        [Browsable(false)]
        public Shape originalShape { get; set; }
    }
}