using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CrystalDecisions.CrystalReports.Engine;
using DinkToPdf;
using DinkToPdf.Contracts;
using HoiThao_Core.Data;
using HoiThao_Core.Data.Repository;
using HoiThao_Core.Helpers;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using System.Text;

namespace HoiThao_Core.Controllers
{
    public class PrintController : Controller
    {
        private readonly IAseanRepository _aseanRepository;

        public PrintController(IAseanRepository aseanRepository)
        {
            _aseanRepository = aseanRepository;
        }
        public IActionResult Index()
        {
            return View();
        }

       

    
    }
}