using System.ComponentModel;

namespace LaRottaO.OfficeTranslationTool.Models
{
    internal class ShapeElement
    {
        [Browsable(false)]
        public int indexOnPresentation { get; set; }

        [Browsable(false)]
        public int slideNumber { get; set; }

        [Browsable(false)]
        public int indexOnSlide { get; set; }

        [Browsable(false)]
        public Boolean belongsToATable { get; set; }

        [Browsable(false)]
        public int parentTableRow { get; set; }

        [Browsable(false)]
        public int parentTableColumn { get; set; }

        [ColumnName("Info")]
        public String info { get; set; }

        [ColumnName("Original Text")]
        public String originalText { get; set; }

        [ColumnName("New Text")]
        public String newText { get; set; }

        [Browsable(false)]
        public Object shape { get; set; }
    }
}