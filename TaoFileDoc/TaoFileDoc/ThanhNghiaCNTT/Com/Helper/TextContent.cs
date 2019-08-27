using TaoFileDoc.ThanhNghiaCNTT.Com.Model;

namespace TaoFileDoc.ThanhNghiaCNTT.Com.Helper
{
    public class TextContent
    {
        /// <summary>
        /// CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM
        /// </summary>
        public string CongHoa = "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM";

        /// <summary>
        /// Độc lập-Tự do-Hạnh phúc
        /// </summary>
        public string DocLap = "Độc lập-Tự do-Hạnh phúc";

        /// <summary>
        /// HỢP ĐỒNG CÔNG VIỆC
        /// </summary>
        public string HopDong = "HỢP ĐỒNG CÔNG VIỆC";

        /// <summary>
        /// Số: {ContractNumber}/KHCN-{TitleCode}
        /// </summary>
        public string SoHopDong = "Số: {ContractNumber}/KHCN-{TitleCode}";

        /// <summary>
        /// Bên A: {FullName}, chủ nhiệm đề tài \"{Title}\", {TitleCode}
        /// </summary>
        public string BeA = "Bên A: {FullName}, chủ nhiệm đề tài \"{Title}\", {TitleCode}";

        /// <summary>
        /// Đơn vi: {WorkUnit}
        /// </summary>
        public string DonViConTac = "Đơn vi: {WorkUnit}";

        /// <summary>
        /// CMND số: {Id}
        /// </summary>
        public string CMND = "CMND số: {Id}";

        /// <summary>
        /// Ngày cấp: {DateId}
        /// </summary>
        public string NgayCap = "Ngày cấp: {DateId}";

        /// <summary>
        /// Nơi cấp:{AddressId}
        /// </summary>
        public string NoiCap = "Nơi cấp:{AddressId}";

        /// <summary>
        /// Mã số thuế: {TaxCode}
        /// </summary>
        public string MaSoThue = "Mã số thuế: {TaxCode}";

        /// <summary>
        /// Điện thoại: {Phone}
        /// </summary>
        public string DienThoai = "Điện thoại: {Phone}";

        /// <summary>
        /// Bên B: CÁC THÀNH VIÊN THAM GIA THỰC HIỆN ĐỀ TÀI GỒM:
        /// </summary>
        public string BenB = "Bên B: CÁC THÀNH VIÊN THAM GIA THỰC HIỆN ĐỀ TÀI GỒM:";

        /// <summary>
        /// {STT}. {FullName} - {Title}
        /// </summary>
        public string ConBenB = "{STT}. {FullName} - {Title}";

        /// <summary>
        /// Function khởi tạo
        /// </summary>
        /// <param name="infoExcel"></param>
        public TextContent(InfoExcel infoExcel)
        {
            SoHopDong = SoHopDong.Replace("{ContractNumber}", infoExcel.ContractNumber).Replace("{TitleCode}", infoExcel.TitleCode);
            BeA = BeA.Replace("{FullName}", infoExcel.A.FullName).Replace("{Title}", infoExcel.A.Title).Replace("{TitleCode}", infoExcel.A.TaxCode);
            DonViConTac = DonViConTac.Replace("{WorkUnit}", infoExcel.A.WorkUnit);
            CMND = CMND.Replace("{Id}", infoExcel.A.Id);
        }
    }
}
