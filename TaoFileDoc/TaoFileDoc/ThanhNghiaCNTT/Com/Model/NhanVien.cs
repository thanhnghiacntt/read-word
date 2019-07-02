using System;

namespace TaoFileDoc.ThanhNghiaCNTT.Com.Model
{
    public class NhanVien
    {
        /// <summary>
        /// Họ và tên
        /// </summary>
        public string FullName { get; set; }

        /// <summary>
        /// Địa chỉ
        /// </summary>
        public string Address { get; set; }

        /// <summary>
        /// Chứng minh nhân dân
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Ngày cấp
        /// </summary>
        public DateTime DateId { get; set; }

        /// <summary>
        /// Nơi cấp
        /// </summary>
        public string AddressId { get; set; }

        /// <summary>
        /// Mã số thuế
        /// </summary>
        public string TaxCode { get; set; }

        /// <summary>
        /// Đơn vị công tác
        /// </summary>
        public string WorkUnit { get; set; }

        /// <summary>
        /// Chức danh
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Hệ số lương
        /// </summary>
        public double CoefficientsSalary { get; set; }

        /// <summary>
        /// Số ngày đã làm việc
        /// </summary>
        public double DayWorked { get; set; }
    }
}
