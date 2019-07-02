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
            string fileName = @"D:\MyProject\VBA\ThongTinConNguoi.docx";
            string fileExcel = @"D:\MyProject\C#\Kiet\Temp\Kiet\Form.xlsx";
            var temp = HelperExcel.GetInfoExcel(fileExcel);
            FileWord fileWord = new FileWord(@"D:\MyProject\C#\Kiet\Data\Mau\CongHoa.docx");
            fileWord.AddContent("Nguyễn Thành Nghĩa \n");
            fileWord.Save(@"D:\abc.doc");
            fileWord.Close();
        }
    }
}
