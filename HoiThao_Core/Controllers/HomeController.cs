using HoiThao_Core.Data;
using HoiThao_Core.Data.Repository;
using HoiThao_Core.Helpers;
using HoiThao_Core.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using Microsoft.AspNetCore;
using DinkToPdf.Contracts;
using DinkToPdf;

namespace HoiThao_Core.Controllers
{
    //    public abstract Microsoft.AspNetCore.Http.QueryString QueryString { get; set; }

    public class HomeController : BaseController
    {

        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly IAseanRepository _aseanRepository;
        private readonly IConverter _converter;

        public HomeController(IHostingEnvironment hostingEnvironment, IAseanRepository aseanRepository, IConverter converter)
        {
            _hostingEnvironment = hostingEnvironment;
            _aseanRepository = aseanRepository;
            _converter = converter;
        }

        private List<OptionListVM> OptionList()
        {
            return new List<OptionListVM>()
            {
                new OptionListVM(){Name = "All", Value = ""},
                new OptionListVM(){Name = "Invited", Value = "true"},
                new OptionListVM(){Name = "Speaker", Value = "false"}
            };
        }
        private List<OptionListVM> OptionEditList()
        {
            return new List<OptionListVM>()
            {
                new OptionListVM(){Name = "-- Selec one --", Value = ""},
                new OptionListVM(){Name = "True", Value = "true"},
                new OptionListVM(){Name = "False", Value = "false"}
            };
        }

        public IActionResult Index(string option, string searchString, int page = 1)
        {
            //string baseUrl = string.Format("{0}://{1}{2}{3}", Request.Scheme, Request.Host, Request.PathBase, Request.QueryString);
            ViewBag.request = UriHelper.GetDisplayUrl(Request);

            ViewData["abc"] = UriHelper.GetDisplayUrl(Request);
            ViewData["CurrentFilter"] = searchString;
            ViewBag.optionFilter = option;

            TempData["optionFilter"] = OptionList();

            var aseans = _aseanRepository.GetAseans(option, searchString, page);

            ViewBag.Aseans = aseans;

            return View(aseans);
        }

        public IActionResult Create(string request)
        {
            ViewBag.request1 = request;
            var a = ViewData["abc"];

            ViewBag.optionFilter = OptionEditList();

            return View();
        }

        [HttpPost]
        public IActionResult Create(Asean asean, string request1)
        {
            
            try
            {
                _aseanRepository.Create(asean);
                SetAlert("Create success.", "success");
            }
            catch (Exception)
            {
                SetAlert("Create success.", "warning");
                throw;
            }

            return Redirect(request1);
        }

        public IActionResult Edit(int id, string request)
        {
            ViewBag.request1 = request;
            var asean = _aseanRepository.GetById(id);

            var editAseanVM = new EditAseanVM()
            {
                Asean = asean,
                OptionListVM = OptionEditList()
            };
            return View(editAseanVM);
        }

        [HttpPost]
        public IActionResult Edit(EditAseanVM editAseanVM, string request1)
        {
            if (ModelState.IsValid)
            {
                _aseanRepository.Update(editAseanVM.Asean);
            }

            return Redirect(request1);
        }

        [HttpDelete]
        public IActionResult Delete(int id)
        {
            var asean = _aseanRepository.GetById(id);
            try
            {
                _aseanRepository.Delete(asean);
                return Json(new
                {
                    status = true
                });
            }
            catch (Exception)
            {
                return Json(new
                {
                    status = false
                });
                throw;
            }
        }

        [HttpPost]
        public IActionResult UpdateCheckin(int id)
        {
            var asean = _aseanRepository.GetById(id);
            asean.Checkin = DateTime.Now;
            try
            {
                _aseanRepository.Update(asean);
                SetAlert("Checkin success.", "success");
            }
            catch (Exception)
            {
                SetAlert("Checkin success.", "warning");
                throw;
            }

            return Json(new
            {
                status = true
            });
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

        public ActionResult PrintReceipt(int id)
        {
            var aseanM = _aseanRepository.GetById(id);
            
            //ViewBag.AmountToString = SoSangChu.DoiSoSangChu(aseanM.Amount.ToString());
            return View(aseanM);
        }

        public ActionResult PrintVAT(int aseanId)
        {
            var aseanM = _aseanRepository.GetById(aseanId);
            string str = aseanM.Amount.ToString();
            if (string.IsNullOrEmpty(str))
                str = "0";

           // ViewBag.AmountToString = SoSangChu.DoiSoSangChu(str);
            return View(aseanM);
        }

        public IActionResult PrintBadge(int id)
        {
            ViewBag.link = UriHelper.GetDisplayUrl(Request);
            var aseanList = _aseanRepository.GetById(id);

            return View(aseanList);
        }

        [HttpGet]
        public IActionResult CreatePDF()
        {
            
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings { Top = 10 },
                DocumentTitle = "PDF Report",
                //Out = @"D:\PDFCreator\Employee_Report.pdf"  USE THIS PROPERTY TO SAVE PDF TO A PROVIDED LOCATION
            };

            var objectSettings = new ObjectSettings
            {
                PagesCount = true,
                //HtmlContent = "<div>abc</div>",
                //Page = link, //USE THIS PROPERTY TO GENERATE PDF CONTENT FROM AN HTML PAGE $"{this.Request.Scheme}://{this.Request.Host}/timeline/Reporting/Show/" + sercret + "/" + configId.ToString(),
                Page = $"{this.Request.Scheme}://{this.Request.Host}",
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "assets", "styles.css") },
                HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Report Footer" }
            };

            var pdf = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
            };

            //_converter.Convert(pdf); IF WE USE Out PROPERTY IN THE GlobalSettings CLASS, THIS IS ENOUGH FOR CONVERSION

            var file = _converter.Convert(pdf);

            //return Ok("Successfully created PDF document.");
            //return File(file, "application/pdf", "EmployeeReport.pdf"); //USE THIS RETURN STATEMENT TO DOWNLOAD GENERATED PDF DOCUMENT
           return File(file, "application/pdf");

        }

    }
}