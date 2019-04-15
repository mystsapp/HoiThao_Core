using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using HoiThao_Core.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Extensions.Configuration;
using HoiThao_Core.Helpers;
using Microsoft.Extensions.DependencyInjection;
using HoiThao_Core.Data;
using HoiThao_Core.Data.Repository;

namespace HoiThao_Core.Controllers
{
    public class HomeController : Controller
    {
        

        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly IAseanRepository _aseanRepository;

        public HomeController(IHostingEnvironment hostingEnvironment, IAseanRepository aseanRepository)
        {
            _hostingEnvironment = hostingEnvironment;
            _aseanRepository = aseanRepository;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult ImportExcel()
        {
            return View();
        }

        public FileResult DownloadExcel()
        {
            string webRootPath = _hostingEnvironment.WebRootPath;
            string folderPath = webRootPath + @"\Doc\";
            string newPath = Path.Combine(webRootPath, folderPath, "Book3.xlsx");

            //return File(newPath, "application/vnd.ms-excel", "Book3.xlsx");

            string filePath = newPath;

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            return File(fileBytes, "application/force-download", "File_mau.xlsx");
        }

        [HttpPost]
        public ActionResult UploadExcel()
        {
            IFormFile file = Request.Form.Files[0];
            string folderName = "excelfolder";
            string webRootPath = _hostingEnvironment.WebRootPath;
            string newPath = Path.Combine(webRootPath, folderName);

            string folderPath = webRootPath + @"\excelfolder\";
            FileInfo fileInfo = new FileInfo(Path.Combine(folderPath, file.FileName));

            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }
            if (file.Length > 0)
            {
                string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                string fullPath = Path.Combine(newPath, file.FileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
               

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"];
                    int totalRows = workSheet.Dimension.Rows;

                    List<Asean> aseanList = new List<Asean>();

                    for (int i = 2; i <= totalRows; i++)
                    {

                        var asean = new Asean();

                        if (workSheet.Cells[i, 2].Value != null)
                            asean.Title = workSheet.Cells[i, 2].Value.ToString();

                        if (workSheet.Cells[i, 3].Value != null)
                            asean.Id = workSheet.Cells[i, 3].Value.ToString();

                        if (workSheet.Cells[i, 4].Value != null)
                            asean.Firstname = workSheet.Cells[i, 4].Value.ToString();

                        if (workSheet.Cells[i, 5].Value != null)
                            asean.Country = workSheet.Cells[i, 5].Value.ToString();

                        if (workSheet.Cells[i, 6].Value != null)
                            asean.Company = workSheet.Cells[i, 6].Value.ToString();

                        if (workSheet.Cells[i, 7].Value != null)
                            asean.Hotel = workSheet.Cells[i, 7].Value.ToString();

                        if (workSheet.Cells[i, 11].Value != null)
                            asean.HotelCheckin = workSheet.Cells[i, 11].Value.ToString();

                        if (workSheet.Cells[i, 12].Value != null)
                            asean.HotelCheckout = workSheet.Cells[i, 12].Value.ToString();

                        if (workSheet.Cells[i, 10].Value != null)
                            asean.HotelPrice = workSheet.Cells[i, 10].Value.ToString();

                        if (workSheet.Cells[i, 17].Value != null)
                            asean.Amount = decimal.Parse(workSheet.Cells[i, 17].Value.ToString());

                        if (workSheet.Cells[i, 18].Value != null)
                            asean.Bankfee = decimal.Parse(workSheet.Cells[i, 18].Value.ToString());

                        if (workSheet.Cells[i, 19].Value != null)
                            asean.Grandtotal = decimal.Parse(workSheet.Cells[i, 19].Value.ToString());

                        if (workSheet.Cells[i, 20].Value != null)
                            asean.Mop = workSheet.Cells[i, 20].Value.ToString();

                        if (workSheet.Cells[i, 23].Value != null)
                            asean.Dfno = workSheet.Cells[i, 23].Value.ToString();

                        if (workSheet.Cells[i, 27].Value != null)
                            asean.Note = workSheet.Cells[i, 27].Value.ToString();

                        if (workSheet.Cells[i, 6].Value != null)
                            asean.Email = workSheet.Cells[i, 6].Value.ToString();

                        if (workSheet.Cells[i, 21].Value != null)
                            asean.Payment = workSheet.Cells[i, 21].Value.ToString();

                        if (workSheet.Cells[i, 21].Value != null)
                            asean.At = workSheet.Cells[i, 21].Value.ToString();

                        if (workSheet.Cells[i, 26].Value != null)
                            asean.Dt = workSheet.Cells[i, 26].Value.ToString();

                        aseanList.Add(asean);

                    }

                    //_db.Customers.AddRange(customerList);
                    //_db.SaveChanges();
                    try
                    {
                        _aseanRepository.AddList(aseanList);

                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                }
            }
            if (System.IO.File.Exists(fileInfo.ToString()))
                System.IO.File.Delete(fileInfo.ToString());

            return Json(new
            {
                status = true
            });
        }

        public ActionResult OnPostImport()
        {
            IFormFile file = Request.Form.Files[0];
            string folderName = "Upload";
            string webRootPath = _hostingEnvironment.WebRootPath;
            string newPath = Path.Combine(webRootPath, folderName);
            StringBuilder sb = new StringBuilder();

            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }
            if (file.Length > 0)
            {
                string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                ISheet sheet;
                string fullPath = Path.Combine(newPath, file.FileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);


                    stream.Position = 0;
                    if (sFileExtension == ".xls")
                    {
                        HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                        sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  
                    }
                    else
                    {
                        XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                        sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook   
                    }
                    IRow headerRow = sheet.GetRow(0); //Get Header Row
                    int cellCount = headerRow.LastCellNum;

                    sb.Append("<table class='table'><tr>");
                    for (int j = 0; j < cellCount; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                        sb.Append("<th>" + cell.ToString() + "</th>");
                    }
                    sb.Append("</tr>");
                    sb.AppendLine("<tr>");
                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                                sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.Append("</table>");

                }
            }

            //DataTable dt = HtmlTableToDataTable.ConvertToDataTable(sb.ToString());

            return this.Content(sb.ToString());
        }
       

    }


}


