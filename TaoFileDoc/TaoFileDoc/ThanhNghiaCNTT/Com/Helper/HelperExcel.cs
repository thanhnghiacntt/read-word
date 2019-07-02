using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using TaoFileDoc.ThanhNghiaCNTT.Com.Model;

namespace TaoFileDoc.ThanhNghiaCNTT.Com.Helper
{
    public class HelperExcel
    {
        public static InfoExcel GetInfoExcel(string pathFile)
        {
            var rs = new InfoExcel()
            {
                B = new List<NhanVien>()
            };
            Application xlApp = new Application
            {
                Visible = false
            };
            object missing = Type.Missing;
            object readOnly = false;
            Workbook xlWorkbook = null;
            try
            {
                xlWorkbook = xlApp.Workbooks.Open(pathFile,
                    missing, readOnly, missing, missing,
                    missing, missing, missing, missing,
                    missing, missing, missing, missing,
                    missing, missing);
                Worksheet mainSheet = xlWorkbook.Sheets["ExportToWord"];
                rs.TitleLevel = GetCell(mainSheet, 3, 4);
                rs.TitleName = GetCell(mainSheet, 4, 4);
                rs.TitleCode = GetCell(mainSheet, 5, 4);
                rs.Content = GetCell(mainSheet, 6, 4);
                rs.ContractNumber = GetCell(mainSheet, 7, 4);
                rs.DateHandoverProduct = ConvertToDateTime(GetCell(mainSheet, 8, 4));
                rs.DateRegisterContract = ConvertToDateTime(GetCell(mainSheet, 9, 4));
                rs.DateRegister = ConvertToDateTime(GetCell(mainSheet, 10, 4));
                rs.DateReceivedProduct = ConvertToDateTime(GetCell(mainSheet, 11, 4));
                rs.DateReceivedMoney = ConvertToDateTime(GetCell(mainSheet, 12, 4));
                var a = new NhanVien()
                {
                    FullName = GetCell(mainSheet, 5, 8),
                    Address = GetCell(mainSheet, 5, 9),
                    Id = GetCell(mainSheet, 5, 10),
                    DateId = ConvertToDateTime(GetCell(mainSheet, 5, 11)),
                    AddressId = GetCell(mainSheet, 5, 12),
                    TaxCode = GetCell(mainSheet, 5, 13),
                    WorkUnit = GetCell(mainSheet, 5, 14)
                };
                for (int i = 9; i < 20; i++)
                {
                    var name = GetCell(mainSheet, i, 8);
                    if (name != null && name.Length > 0)
                    {
                        var b = new NhanVien()
                        {
                            FullName = name,
                            Address = GetCell(mainSheet, i, 9),
                            Id = GetCell(mainSheet, i, 10),
                            DateId = ConvertToDateTime(GetCell(mainSheet, i, 11)),
                            AddressId = GetCell(mainSheet, i, 12),
                            TaxCode = GetCell(mainSheet, i, 13),
                            WorkUnit = GetCell(mainSheet, i, 14),
                            Title = GetCell(mainSheet, i, 15),
                            CoefficientsSalary = double.Parse(GetCell(mainSheet, i, 16)),
                            DayWorked = double.Parse(GetCell(mainSheet, i, 17)),
                        };
                        rs.B.Add(b);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
            }
            xlApp.Quit();
            return rs;
        }

        /// <summary>
        /// Get value
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <returns></returns>
        public static string GetCell(Worksheet sheet, int i, int j)
        {
            if (sheet.Cells != null && sheet.Cells[i, j] != null && sheet.Cells[i, j].Value != null)
            {
                var cellValue = (sheet.Cells[i, j] as Range).Value.ToString();
                return cellValue;
            }
            return null;
        }

        /// <summary>
        /// Convert string to datetime
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static DateTime ConvertToDateTime(string str)
        {
            DateTime rs = DateTime.Now;
            if (str != null)
            {
                str = str.Substring(0, 10);
                rs = DateTime.ParseExact(str, "dd/MM/yyyy", null);
            }
            return rs;
        }
    }
}
