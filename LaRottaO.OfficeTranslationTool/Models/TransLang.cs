using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace LaRottaO.OfficeTranslationTool.Models
{
    public class TransLang
    {
   

        public String languageCode { get; set; }

        public Microsoft.Office.Core.MsoLanguageID officeLanguageId { get; set; }

        public TransLang( string languageCode, MsoLanguageID officeLanguageId)
        {       
            this.languageCode = languageCode;
            this.officeLanguageId = officeLanguageId;
        }
    }


}
