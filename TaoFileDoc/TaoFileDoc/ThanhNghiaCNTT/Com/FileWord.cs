using Microsoft.Office.Interop.Word;
using System;
using Application = Microsoft.Office.Interop.Word.Application;

namespace TaoFileDoc.ThanhNghiaCNTT.Com
{
    public class FileWord
    {
        public Document Document { get; private set; }
        public Application Winword { get; private set; }

        public string Path { get; set; }

        public FileWord(string path)
        {
            Path = path;
            //Create an instance for word app  
            Winword = new Application
            {
                //Set animation status for word application  
                ShowAnimation = false,
                //Set status for word application is to be visible or not.  
                Visible = false
            };
            Document = Winword.Documents.Open(Path);
        }

        /// <summary>
        /// Close
        /// </summary>
        public void Close()
        {
            if (Document != null)
            {
                Document.Close();
                Winword.Quit();
            }
        }

        /// <summary>
        /// Save file
        /// </summary>
        /// <param name="path"></param>
        public void Save(string path)
        {
            if (Document != null)
            {
                Document.SaveAs2(path);
            }
        }

        public void AddContent(string content)
        {
            Document.Content.InsertAfter(content);
            Document.Content.Select();
            Document.Content.set_Style(Document.Styles["Title"]);
            //Document.Paragraphs.Add(range);
        }

        private void InsertMultiFormatParagraph(string psText, int piSize, int piSpaceAfter = 10)
        {
            object mobjMissing = null;
            object start = Document.Content.Start;
            object end = Document.Content.End;

            Document.Range(ref start, ref end).Select();

            Paragraph para = Document.Content.Paragraphs.Add(ref mobjMissing);

            para.Range.Text = psText;
            // Explicitly set this to "not bold"
            para.Range.Font.Bold = 0;
            para.Range.Font.Size = piSize;
            para.Format.SpaceAfter = piSpaceAfter;

            object objStart = para.Range.Start;
            object objEnd = para.Range.Start + psText.IndexOf(":");

            Range rngBold = Document.Range(ref objStart, ref objEnd);
            rngBold.Bold = 1;

            para.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// Create content
        /// </summary>
        /// <param name="str"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public Range CreateContent(string str, int start = 0, int end = 0)
        {
            Document.Content.SetRange(start, end);
            Document.Content.Text = str + Environment.NewLine;
            return Document.Content;
        }
    }
}
