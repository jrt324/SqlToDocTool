using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XWPF.UserModel;

namespace CommonService
{
    public static class ExtensionMethods
    {
        public static XWPFParagraph CreateParagraph(this XWPFTableCell aCell)
        {
            XWPFParagraph pIO = aCell.Paragraphs.Count > 0 ? aCell.Paragraphs[0] : aCell.AddParagraph();
            pIO.Alignment = ParagraphAlignment.LEFT;
            return pIO;
        }
    }
}
