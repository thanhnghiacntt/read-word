using System;
using System.Collections.Generic;

namespace TaoFileDoc.ThanhNghiaCNTT.Com.Model
{
    public class InfoExcel
    {
        /// <summary>
        /// Bên A
        /// </summary>
        public NhanVien A { get; set; }

        /// <summary>
        /// Bên B
        /// </summary>
        public IList<NhanVien> B { get; set; }

        /// <summary>
        /// Cấp đề tài
        /// </summary>
        public string TitleLevel { get; set; }

        /// <summary>
        /// Tên đề tài
        /// </summary>
        public string TitleName { get; set; }

        /// <summary>
        /// Mã đề tài
        /// </summary>
        public string TitleCode { get; set; }

        /// <summary>
        /// Nội dung công việc
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// Số hợp đồng
        /// </summary>
        public string ContractNumber { get; set; }

        /// <summary>
        /// Ngày bàn giao sản phẩm
        /// </summary>
        public DateTime DateHandoverProduct { get; set; }

        /// <summary>
        /// Ngày ký hợp đồng
        /// </summary>
        public DateTime DateRegisterContract { get; set; }

        /// <summary>
        /// Ngày ký biên bản
        /// </summary>
        public DateTime DateRegister { get; set; }

        /// <summary>
        /// Ngày bàn giao sản phẩm
        /// </summary>
        public DateTime DateReceivedProduct { get; set; }

        /// <summary>
        /// Ngày nhận tiền
        /// </summary>
        public DateTime DateReceivedMoney { get; set; }
    }
}
