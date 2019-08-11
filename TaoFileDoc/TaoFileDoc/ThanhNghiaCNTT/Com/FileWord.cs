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
                Visible = true
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

        /// <summary>
        /// Add content
        /// </summary>
        /// <param name="content"></param>
        /// <param name="style"></param>
        public void AddContent(string content, string style = null)
        {
            Document.Content.InsertAfter(content);
            Document.Content.Select();
            if (style != null)
            {
                Document.Content.set_Style(Document.Styles[style]);
            }
        }

        /// <summary>
        /// Add content
        /// </summary>
        /// <param name="content"></param>
        /// <param name="style"></param>
        public void AddContentNewLine(string content, string style = null)
        {
            AddContent("\n" + content, style);
        }
    }
}
