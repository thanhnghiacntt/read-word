using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TaoFileDoc.ThanhNghiaCNTT.Com;
using TaoFileDoc.ThanhNghiaCNTT.Com.Helper;
using TaoFileDoc.ThanhNghiaCNTT.Com.Model;

namespace TaoFileDoc
{
    public partial class Main : Form
    {

        public Main()
        {
            InitializeComponent();
        }

        private void BntStart_Click(object sender, EventArgs e)
        {
            string fileName = @"D:\MyProjects\VBA\ThongTinConNguoi.docx";
            string fileExcel = @"D:\MyProjects\C#\Kiet\Temp\Kiet\Form.xlsx";
            var temp = HelperExcel.GetInfoExcel(fileExcel);
            AddFile(temp, @"D:\MyProjects\C#\read-word\Data\mau.docx");
        }

        private void AddFile(InfoExcel infoExcel, string path)
        {

            FileWord fileWord = new FileWord(path);
            fileWord.AddContent(TextContent.CongHoa, "Heading 1");
            fileWord.AddContentNewLine(TextContent.DocLap, "Heading 2");
            fileWord.AddContentNewLine(TextContent.HopDong);
            fileWord.AddContentNewLine(TextContent.SoHopDong);
            fileWord.Save(@"D:\abc.doc");
            fileWord.Close();
        }
    }
}
