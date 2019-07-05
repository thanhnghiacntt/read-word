using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TaoFileDoc.ThanhNghiaCNTT.Com.Helper
{
    public class HelperWord
    {
        public static void FindAndReplace(Document document, object findText, object replaceWithText)
        {
            object o = Missing.Value;
            object oFalse = false;
            object oTrue = true;
            foreach (Range range in document.StoryRanges)
            {
                Find find = range.Find;
                object replace = WdReplace.wdReplaceAll;
                object findWrap = WdFindWrap.wdFindContinue;
                find.Execute(ref findText, ref o, ref o, ref o, ref oFalse, ref o,
                    ref o, ref findWrap, ref o, ref replaceWithText,
                    ref replace, ref o, ref o, ref o, ref o);
                Marshal.FinalReleaseComObject(find);
                Marshal.FinalReleaseComObject(range);
            }
        }

        public static Paragraph InsertedText(Document document, string text)
        {
            object missing = Type.Missing;
            var rs = document.Content.Paragraphs.Add(ref missing);
            rs.Range.Text = text;
            rs.Range.InsertParagraphAfter();
            return rs;
        }
    }
}
