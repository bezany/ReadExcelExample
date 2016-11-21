using ExcelExampleData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;

namespace ExcelExampleWebApp.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Index()
        {
            return View();
        }

    
        [HttpPost]
        public JsonResult GetExcelProviders()
        {
            var providers = new List<string>(new []{"Test", "test2"});
            return Json(providers);
        }

        [HttpPost]
        public JsonResult SearchUseNPOI(string searchName)
        {
            var data = new NPOIProvider(Server.MapPath("~/App_Data/test.xlsx")).SearchRow(searchName);
            return Json(data);
        }

        [HttpPost]
        public JsonResult SearchUseEPPlus(string searchName)
        {
            var data = new EPPlusProvier(Server.MapPath("~/App_Data/test.xlsx")).SearchRow(searchName);
            return Json(data);
        }
    }
}