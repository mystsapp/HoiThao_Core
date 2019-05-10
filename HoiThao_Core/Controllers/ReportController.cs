using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Data.Repository;
//using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml.Style;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HoiThao_Core.Controllers
{
    public class ReportController : Microsoft.AspNetCore.Mvc.Controller
    {
        private readonly IAseanRepository _aseanRepository;
        private readonly IHostingEnvironment _hostingEnvironment;

        public ReportController(IAseanRepository aseanRepository, IHostingEnvironment hostingEnvironment)
        {
            _aseanRepository = aseanRepository;
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }

        //public async Task<IActionResult> OnPostExport()
        //{
        //    string sWebRootFolder = _hostingEnvironment.WebRootPath;
        //    string sFileName = @"demo.xlsx";
        //    string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
        //    FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
        //    var memory = new MemoryStream();
        //    using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
        //    {
        //        IWorkbook workbook;
        //        workbook = new XSSFWorkbook();
        //        ISheet excelSheet = workbook.CreateSheet("Demo");
        //        IRow row = excelSheet.CreateRow(0);

        //        row.CreateCell(0).SetCellValue("ID");
        //        row.CreateCell(1).SetCellValue("Name");
        //        row.CreateCell(2).SetCellValue("Age");

        //        row = excelSheet.CreateRow(1);
        //        row.CreateCell(0).SetCellValue(1);
        //        row.CreateCell(1).SetCellValue("Kane Williamson");
        //        row.CreateCell(2).SetCellValue(29);

        //        row = excelSheet.CreateRow(2);
        //        row.CreateCell(0).SetCellValue(2);
        //        row.CreateCell(1).SetCellValue("Martin Guptil");
        //        row.CreateCell(2).SetCellValue(33);

        //        row = excelSheet.CreateRow(3);
        //        row.CreateCell(0).SetCellValue(3);
        //        row.CreateCell(1).SetCellValue("Colin Munro");
        //        row.CreateCell(2).SetCellValue(23);

        //        workbook.Write(fs);
        //    }
        //    using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
        //    {
        //        await stream.CopyToAsync(memory);
        //    }
        //    memory.Position = 0;
        //    return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);
        //}

        public IActionResult ConferenceReport()
        {
            //string rootFolder = _hostingEnvironment.WebRootPath;
            //string fileName = "CONFERENCE_GROUP_BY_COUNTRY_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx";

            //FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            byte[] fileContents;    

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 40;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Country
            xlSheet.Column(6).Width = 10;// Telephone
            xlSheet.Column(7).Width = 25;//Email

            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 25;// Total A
            xlSheet.Column(10).Width = 40;// Total B
            xlSheet.Column(11).Width = 20;// Grand Total
            xlSheet.Column(12).Width = 20;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CONFERENCE REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.ConferenceReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.Headers.Add("Content-Disposition", "attachment; filename=" + "CONFERENCE_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();

                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CONFERENCE_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx"
            );

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult WorkshopReport()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 25;// Total A
            xlSheet.Column(10).Width = 40;// Total B
            xlSheet.Column(11).Width = 20;// Grand Total
            xlSheet.Column(12).Width = 20;// Paid
            xlSheet.Column(13).Width = 10;// Currency

            xlSheet.Column(14).Width = 15;//Lectures only

            xlSheet.Column(15).Width = 15;// Lectures & Hands-on Cadaveric Dissection
            xlSheet.Column(16).Width = 25;//Lectures only
            xlSheet.Column(17).Width = 25;//Lecture & Live Surgery Demonstration
            xlSheet.Column(18).Width = 25;//Lectures only
            xlSheet.Column(19).Width = 25;//Lecture & Live Surgery Demonstration
            xlSheet.Column(20).Width = 25;//4. Allergy & Immunology


            xlSheet.Cells[2, 1].Value = "WORKSHOP REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;



            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Department";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";

            xlSheet.Cells[4, 14, 4, 15].Merge = true;
            xlSheet.Cells[4, 14].Value = "1. Functional Endoscopic Sinus Surgery";
            xlSheet.Row(4).Height = MeasureTextHeight("1. Functional Endoscopic Sinus Surgery", xlSheet.Cells[4, 14].Style.Font, 25);
            xlSheet.Cells[4, 14].Style.WrapText = true;
            xlSheet.Cells[5, 14].Value = "Lectures only";

            xlSheet.Cells[5, 15].Value = "Lectures & Hands-on Cadaveric Dissection";
            xlSheet.Cells[5, 15].Style.WrapText = true;

            xlSheet.Cells[4, 16, 4, 17].Merge = true;
            xlSheet.Cells[4, 16].Value = "2. Plastic Surgery";
            xlSheet.Cells[5, 17].Style.WrapText = true;
            // xlSheet.Row(4).Height = MeasureTextHeight("2. Plastic Surgery", xlSheet.Cells[4, 16].Style.Font, 25);
            xlSheet.Cells[5, 16].Value = "Lectures only";
            xlSheet.Cells[5, 17].Value = "Lecture & Live Surgery Demonstration";

            xlSheet.Cells[4, 18, 4, 19].Merge = true;
            xlSheet.Cells[4, 18].Value = "3. Sleep Disordered Breathing";
            xlSheet.Cells[5, 19].Style.WrapText = true;
            //xlSheet.Row(4).Height = MeasureTextHeight("3. Sleep Disordered Breathing", xlSheet.Cells[4, 18].Style.Font, 25);

            xlSheet.Cells[5, 18].Value = "Lectures only";
            xlSheet.Cells[5, 19].Value = "Lecture & Live Surgery Demonstration";
            xlSheet.Cells[5, 20].Value = "4. Allergy & Immunology";
            xlSheet.Cells[5, 1, 5, 20].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.ConferenceReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(4, 1, 4 + dt.Rows.Count, 20, xlSheet);
            setFontBold(4, 1, 4, 20, 12, xlSheet);
            setFontBold(5, 1, 5, 20, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 20, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 20, xlSheet);
            setFontBold(dong, 5, dong, 20, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "WORKSHOP_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CONFERENCE_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult ConferenceGroupByCountry()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Total A
            xlSheet.Column(10).Width = 15;// Total B
            xlSheet.Column(11).Width = 15;// Grand Total
            xlSheet.Column(12).Width = 15;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CONFERENCE GROUP BY COUNTRY REPORT ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.ConferenceGroupByCountry();

            var countryList = _aseanRepository.GetAllCountry();
            var count = countryList.Count();
            if (dt != null)
            {
                foreach (var item in countryList)
                {
                    dong++;
                    var listItem = _aseanRepository.GetAllByCountry(item);
                    dong++;
                    xlSheet.Cells[dong, 5].Value = item;
                    //setBorder(dong, 1, dong, 16, xlSheet);
                    //xlSheet.Cells[dong, 1, dong, 16].Merge = true;
                    setCenterAligment(dong, 1, dong, 7, xlSheet);

                    dong++;
                    for (int i = 0; i < listItem.Rows.Count; i++)
                    {


                        dong++;
                        for (int j = 0; j < listItem.Columns.Count; j++)
                        {
                            setBorder(dong - 2, 1, dong, 16, xlSheet);
                            if (String.IsNullOrEmpty(listItem.Rows[i][j].ToString()))
                            {
                                xlSheet.Cells[dong, j + 1].Value = "";
                            }
                            else
                            {
                                xlSheet.Cells[dong, j + 1].Value = listItem.Rows[i][j];
                            }


                        }
                    }
                }

            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            //setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);


            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "CONFERENCE_GROUP_BY_COUNTRY_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();

                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CONFERENCE_GROUP_BY_COUNTRY_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult InvitedGuests()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Total A
            xlSheet.Column(10).Width = 15;// Total B
            xlSheet.Column(11).Width = 15;// Grand Total
            xlSheet.Column(12).Width = 15;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CONFERENCE REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            //DataTable dt = _aseanService.ConferenceReport();
            DataTable dt = _aseanRepository.ConferenceGroupByCountry();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "INVITED_GUEST_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();

                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "INVITED_GUEST_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");


            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult Received()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Total A
            xlSheet.Column(10).Width = 15;// Total B
            xlSheet.Column(11).Width = 15;// Grand Total
            xlSheet.Column(12).Width = 15;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Column(17).Width = 25;//Invited
            xlSheet.Column(18).Width = 10;//Amount
            xlSheet.Column(19).Width = 15;// Country
            xlSheet.Column(20).Width = 15;// Pick up
            xlSheet.Column(21).Width = 15;// See off
            xlSheet.Column(22).Width = 15;// Tour
            xlSheet.Column(23).Width = 10;// Hotel
            xlSheet.Column(24).Width = 25;//Note

            xlSheet.Cells[2, 1].Value = "PAYMENT REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total ";
            xlSheet.Cells[5, 12].Value = "Paid ";
            xlSheet.Cells[5, 13].Value = "Currency ";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident ";
            xlSheet.Cells[5, 16].Value = "Accompany Persons ";

            xlSheet.Cells[5, 17].Value = "Invited";
            xlSheet.Cells[5, 18].Value = "Amount";
            xlSheet.Cells[5, 19].Value = "Country ";
            xlSheet.Cells[5, 20].Value = "Pick up ";
            xlSheet.Cells[5, 21].Value = "See off ";
            xlSheet.Cells[5, 22].Value = "Tour";
            xlSheet.Cells[5, 23].Value = "Hotel";
            xlSheet.Cells[5, 24].Value = "Note";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.PaymentReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 24, xlSheet);
            setFontBold(5, 1, 5, 24, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 24, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 24, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "PAYMENT_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "PAYMENT_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult PickupAtTheAirport()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Total A
            xlSheet.Column(10).Width = 15;// Total B
            xlSheet.Column(11).Width = 15;// Grand Total
            xlSheet.Column(12).Width = 15;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Bus

            xlSheet.Column(15).Width = 10;// 4 seat car
            xlSheet.Column(16).Width = 25;//Arrival

            xlSheet.Column(17).Width = 25;//Flight No
            xlSheet.Column(18).Width = 10;//Accompany person

            xlSheet.Cells[2, 1].Value = "PICKUP REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total ";
            xlSheet.Cells[5, 12].Value = "Paid ";
            xlSheet.Cells[5, 13].Value = "Currency ";
            xlSheet.Cells[5, 14].Value = "Bus ";

            xlSheet.Cells[5, 15].Value = "4 seat car ";
            xlSheet.Cells[5, 16].Value = "Arrival ";

            xlSheet.Cells[5, 17].Value = "Flight No";
            xlSheet.Cells[5, 18].Value = "Accompany person";

            xlSheet.Cells[5, 1, 5, 18].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.PickupReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            setBorder(5, 1, 5 + dt.Rows.Count, 18, xlSheet);
            setFontBold(5, 1, 5, 18, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 18, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 18, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "PICKUP_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "PICKUP_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult SeeOffAtTheAirport()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Total A
            xlSheet.Column(10).Width = 15;// Total B
            xlSheet.Column(11).Width = 15;// Grand Total
            xlSheet.Column(12).Width = 15;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Bus

            xlSheet.Column(15).Width = 10;// 4 seat car
            xlSheet.Column(16).Width = 25;//Arrival

            xlSheet.Column(17).Width = 25;//Flight No
            xlSheet.Column(18).Width = 10;//Accompany person

            xlSheet.Cells[2, 1].Value = "See Off At The Air REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total ";
            xlSheet.Cells[5, 12].Value = "Paid ";
            xlSheet.Cells[5, 13].Value = "Currency ";
            xlSheet.Cells[5, 14].Value = "Bus ";

            xlSheet.Cells[5, 15].Value = "4 seat car ";
            xlSheet.Cells[5, 16].Value = "Arrival ";

            xlSheet.Cells[5, 17].Value = "Flight No";
            xlSheet.Cells[5, 18].Value = "Accompany person";

            xlSheet.Cells[5, 1, 5, 18].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.PickupReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            setBorder(5, 1, 5 + dt.Rows.Count, 18, xlSheet);
            setFontBold(5, 1, 5, 18, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 18, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 18, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "SEE_OFF_AT_THE_AIR_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "SEE_OFF_AT_THE_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult BuyAirTicket()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Y Class
            xlSheet.Column(10).Width = 15;// C Class

            xlSheet.Cells[2, 1].Value = "AIRTICKET REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Y Class ";
            xlSheet.Cells[5, 10].Value = "C Class ";

            xlSheet.Cells[5, 1, 5, 18].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.AirticketReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 10, xlSheet);
            setFontBold(5, 1, 5, 10, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 10, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 10, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "AIRTICKET_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "AIRTICKET_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult Tour()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 15;// Pax
            xlSheet.Column(10).Width = 15;// Hotel

            xlSheet.Cells[2, 1].Value = "TOUR REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Pax ";
            xlSheet.Cells[5, 10].Value = "Hotel ";

            xlSheet.Cells[5, 1, 5, 18].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.TourReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "failure"
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 10, xlSheet);
            setFontBold(5, 1, 5, 10, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 10, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 10, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "TOUR_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "TOUR_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        [HttpPost]
        public ActionResult ExportHotel(string hotel)
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Registration date
            xlSheet.Column(3).Width = 30;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Department
            xlSheet.Column(6).Width = 20;// Telephone
            xlSheet.Column(7).Width = 25;//Email
            xlSheet.Column(8).Width = 10;//ID

            xlSheet.Column(9).Width = 10;// Checkin
            xlSheet.Column(10).Width = 15;// Checkout
            xlSheet.Column(11).Width = 15;// Rate
            xlSheet.Column(12).Width = 10;// Booking inf
            xlSheet.Column(13).Width = 15;// Group

            xlSheet.Cells[2, 1].Value = hotel.ToUpper() + " REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Registration date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Department";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Checkin ";
            xlSheet.Cells[5, 10].Value = "Checkout ";

            xlSheet.Cells[5, 11].Value = "Rate";
            xlSheet.Cells[5, 12].Value = "Booking inf ";
            xlSheet.Cells[5, 13].Value = "Group ";

            xlSheet.Cells[5, 1, 5, 13].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.HotelReport(hotel);


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "Export failure."
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 13, xlSheet);
            setFontBold(5, 1, 5, 13, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 13, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            //setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 13, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + hotel.ToUpper() + "_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: hotel.ToUpper() + "_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return Json(new
            {
                status = false,
                message = "Export success."
            });
        }

        public ActionResult CheckinList()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Check in date
            xlSheet.Column(3).Width = 40;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Country
            xlSheet.Column(6).Width = 10;// Telephone
            xlSheet.Column(7).Width = 25;//Email

            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 25;// Total A
            xlSheet.Column(10).Width = 40;// Total B
            xlSheet.Column(11).Width = 20;// Grand Total
            xlSheet.Column(12).Width = 20;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CHECKIN REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Check in date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.CheckinReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "Checkin is null."
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult VatBill()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(3).Width = 40;// Name
            xlSheet.Column(4).Width = 20;// Company
            xlSheet.Column(5).Width = 20;// Description
            xlSheet.Column(6).Width = 10;// VAT Date
            xlSheet.Column(7).Width = 10;// Curr.
            xlSheet.Column(8).Width = 25;// VAT Bill

            xlSheet.Column(9).Width = 25;// Rate
            xlSheet.Column(10).Width = 40;// Revenue
            xlSheet.Column(11).Width = 20;// Banking fee
            xlSheet.Column(12).Width = 20;// VAT
            xlSheet.Column(13).Width = 25;// Total

            xlSheet.Column(14).Width = 10;// Tax code

            xlSheet.Cells[2, 1].Value = "VAT REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT ";
            xlSheet.Cells[5, 2].Value = "ID ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Company ";
            xlSheet.Cells[5, 5].Value = "Description";
            xlSheet.Cells[5, 6].Value = "VAT Date";
            xlSheet.Cells[5, 7].Value = "Curr.";

            xlSheet.Cells[5, 8].Value = "VAT Bill";
            xlSheet.Cells[5, 9].Value = "Rate ";
            xlSheet.Cells[5, 10].Value = "Revenue ";
            xlSheet.Cells[5, 11].Value = "Banking fee ";
            xlSheet.Cells[5, 12].Value = "VAT";
            xlSheet.Cells[5, 13].Value = "Total";
            xlSheet.Cells[5, 14].Value = "Tax code";

            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.VatReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "Vat is null."
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 14, xlSheet);
            setFontBold(5, 1, 5, 14, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 14, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 14, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "VAT_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "VAT_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult OnSiteList()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Check in date
            xlSheet.Column(3).Width = 40;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Country
            xlSheet.Column(6).Width = 10;// Telephone
            xlSheet.Column(7).Width = 25;//Email

            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 25;// Total A
            xlSheet.Column(10).Width = 40;// Total B
            xlSheet.Column(11).Width = 20;// Grand Total
            xlSheet.Column(12).Width = 20;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CHECKIN REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Check in date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.CheckinReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "Checkin is null."
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }

        public ActionResult Speaker()
        {
            byte[] fileContents;

            ExcelPackage ExcelApp = new ExcelPackage();
            ExcelWorksheet xlSheet = ExcelApp.Workbook.Worksheets.Add("Report");
            // Định dạng chiều dài cho cột
            xlSheet.Column(1).Width = 10;//STT
            xlSheet.Column(2).Width = 25;// Check in date
            xlSheet.Column(3).Width = 40;// Name
            xlSheet.Column(4).Width = 20;// Address
            xlSheet.Column(5).Width = 20;// Country
            xlSheet.Column(6).Width = 10;// Telephone
            xlSheet.Column(7).Width = 25;//Email

            xlSheet.Column(8).Width = 10;//ID
            xlSheet.Column(9).Width = 25;// Total A
            xlSheet.Column(10).Width = 40;// Total B
            xlSheet.Column(11).Width = 20;// Grand Total
            xlSheet.Column(12).Width = 20;// Paid
            xlSheet.Column(13).Width = 10;// Currency
            xlSheet.Column(14).Width = 25;//Full delegate

            xlSheet.Column(15).Width = 10;// Resident
            xlSheet.Column(16).Width = 25;//Accompany Persons

            xlSheet.Cells[2, 1].Value = "CHECKIN REPORT  ";
            xlSheet.Cells[2, 1].Style.Font.SetFromFont(new Font("Times New Roman", 16, FontStyle.Bold));
            xlSheet.Cells[2, 1, 2, 7].Merge = true;
            setCenterAligment(2, 1, 2, 7, xlSheet);
            // dinh dang tu ngay den ngay
            //if (tungay == denngay)
            //{
            //    fromTo = "Ngày: " + tungay;
            //}
            //else
            //{
            //    fromTo = "Từ ngày: " + tungay + " đến ngày: " + denngay;
            //}

            //xlSheet.Cells[3, 1].Value = fromTo;
            //xlSheet.Cells[3, 1, 3, 7].Merge = true;
            //xlSheet.Cells[3, 1].Style.Font.SetFromFont(new Font("Times New Roman", 14, FontStyle.Bold));
            //setCenterAligment(3, 1, 3, 7, xlSheet);

            // Tạo header
            xlSheet.Cells[5, 1].Value = "STT";
            xlSheet.Cells[5, 2].Value = "Check in date ";
            xlSheet.Cells[5, 3].Value = "Name ";
            xlSheet.Cells[5, 4].Value = "Address";
            xlSheet.Cells[5, 5].Value = "Country";
            xlSheet.Cells[5, 6].Value = "Telephone";
            xlSheet.Cells[5, 7].Value = "Email";

            xlSheet.Cells[5, 8].Value = "ID";
            xlSheet.Cells[5, 9].Value = "Total A ";
            xlSheet.Cells[5, 10].Value = "Total B ";
            xlSheet.Cells[5, 11].Value = "Grand Total";
            xlSheet.Cells[5, 12].Value = "Paid";
            xlSheet.Cells[5, 13].Value = "Currency";
            xlSheet.Cells[5, 14].Value = "Full delegate";

            xlSheet.Cells[5, 15].Value = "Resident";
            xlSheet.Cells[5, 16].Value = "Accompany Persons";
            xlSheet.Cells[5, 1, 5, 16].Style.Font.SetFromFont(new Font("Times New Roman", 12, FontStyle.Bold));

            // do du lieu tu table
            int dong = 5;


            DataTable dt = _aseanRepository.SpeakerReport();


            if (dt != null)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dong++;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                        {
                            xlSheet.Cells[dong, j + 1].Value = "";
                        }
                        else
                        {
                            xlSheet.Cells[dong, j + 1].Value = dt.Rows[i][j];
                        }
                    }
                }
            }
            else
            {
                // SetAlert("No sale.", "warning");
                return Json(new
                {
                    status = false,
                    message = "Speaker is null."
                });
            }
            dong++;
            //// Merger cot 4,5 ghi tổng tiền
            //setRightAligment(dong, 3, dong, 3, xlSheet);
            //xlSheet.Cells[dong, 1, dong, 2].Merge = true;
            //xlSheet.Cells[dong, 1].Value = "Tổng tiền: ";

            // Sum tổng tiền
            xlSheet.Cells[dong, 5].Value = "TOTAL";
            xlSheet.Cells[dong, 11].Formula = "SUM(K6:K" + (6 + dt.Rows.Count - 1) + ")";

            //xlSheet.Cells[dong, 5].Value = "TC";
            //xlSheet.Cells[dong, 6].Formula = "SUM(F6:F" + (6 + dt.Rows.Count - 1) + ")";
            //xlSheet.Cells[dong, 7].Formula = "SUM(G6:G" + (6 + dt.Rows.Count - 1) + ")";

            // định dạng số

            //NumberFormat(dong, 6, dong, 7, xlSheet);

            setBorder(5, 1, 5 + dt.Rows.Count, 16, xlSheet);
            setFontBold(5, 1, 5, 16, 12, xlSheet);
            setFontSize(6, 1, 6 + dt.Rows.Count, 16, 12, xlSheet);
            // dinh dang giua cho cot stt
            setCenterAligment(6, 1, 6 + dt.Rows.Count, 1, xlSheet);

            setBorder(dong, 5, dong, 16, xlSheet);
            setFontBold(dong, 5, dong, 16, 12, xlSheet);

            // dinh dạng ngay thang cho cot ngay di , ngay ve
            DateTimeFormat(6, 2, 6 + dt.Rows.Count, 2, xlSheet);
            // canh giưa cot  ngay di, ngay ve, so khach 
            setCenterAligment(6, 4, 6 + dt.Rows.Count, 6, xlSheet);
            // dinh dạng number cot doanh so
            NumberFormat(6, 11, 6 + dt.Rows.Count, 11, xlSheet);
            NumberFormat(6, 12, 6 + dt.Rows.Count, 12, xlSheet);

            //xlSheet.View.FreezePanes(6, 20);

            try
            {
                //Response.Clear();
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");
                //Response.BinaryWrite(ExcelApp.GetAsByteArray());
                //Response.End();
                fileContents = ExcelApp.GetAsByteArray();
                return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "CHECKIN_REPORT" + System.DateTime.Now.ToString("dd/MM/yyyy HH:mm") + ".xlsx");

            }
            catch (Exception)
            {

                throw;
            }

            return View();
        }


        /////////////////////////////////////////////////////////////////////////////////////////////////////////


        private static void NumberFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                range.Style.Numberformat.Format = "#,#0";
            }
        }
        private static void DateFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Numberformat.Format = "dd/MM/yyyy";
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void DateTimeFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void setRightAligment(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
        }
        private static void setCenterAligment(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
        private static void setFontSize(int fromRow, int fromColumn, int toRow, int toColumn, int fontSize, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Font.SetFromFont(new Font("Times New Roman", fontSize, FontStyle.Regular));
            }
        }
        private static void setFontBold(int fromRow, int fromColumn, int toRow, int toColumn, int fontSize, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Font.SetFromFont(new Font("Times New Roman", fontSize, FontStyle.Bold));
            }
        }
        private static void setBorder(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
        }
        private static void PhantramFormat(int fromRow, int fromColumn, int toRow, int toColumn, ExcelWorksheet sheet)
        {
            using (var range = sheet.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Numberformat.Format = "0 %";
            }
        }

        public double MeasureTextHeight(string text, ExcelFont font, int width)
        {
            if (string.IsNullOrEmpty(text)) return 0.0;
            var bitmap = new Bitmap(1, 1);
            var graphics = Graphics.FromImage(bitmap);

            var pixelWidth = Convert.ToInt32(width * 7.5);  //7.5 pixels per excel column width
            var drawingFont = new Font(font.Name, font.Size);
            var size = graphics.MeasureString(text, drawingFont, pixelWidth);

            //72 DPI and 96 points per inch.  Excel height in points with max of 409 per Excel requirements.
            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96, 409);
        }
    }
}