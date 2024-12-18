using Microsoft.Office.Interop.PowerPoint;
using System;
using System.ComponentModel;
using static LaRottaO.OfficeTranslationTool.GlobalConstants;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Models
{
    public class PptShape
    {
        [Browsable(false)]
        public int indexOnPresentation { get; set; }

        [Browsable(false)]
        public int slideNumber { get; set; }

        [Browsable(false)]
        public int indexOnSlide { get; set; }

        //[Browsable(false)]
        public dynamic section { get; set; }

        //[Browsable(false)]
        public dynamic internalId { get; set; }

        [Browsable(false)]
        public Boolean belongsToATable { get; set; }

        [Browsable(false)]
        public int parentTableRow { get; set; }

        [Browsable(false)]
        public int parentTableColumn { get; set; }

        [ColumnName("Info")]
        public String info { get; set; } = String.Empty;

        //[Browsable(false)]
        [ColumnName("Element type")]
        public ElementType type { get; set; }

        [ColumnName("Original Text")]
        public String originalText { get; set; } = String.Empty;

        [ColumnName("New Text")]
        public String newText { get; set; } = String.Empty;

        [Browsable(false)]
        public Shape originalShape { get; set; }
    }
}