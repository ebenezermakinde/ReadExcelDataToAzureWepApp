using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult Excels()
        {
            ViewBag.Message = "the data from excel:";
            string data = "";

            //your excel path after uploaded, here I hardcoded it for test only
            string path = @"D:\home\site\wwwroot\Files\ddd.xlsx";
            var excelData = new ExcelData(path);
            var people = excelData.GetData("sheet1");

            foreach (var p in people)
            {
                for (int i = 0; i <= p.ItemArray.GetUpperBound(0); i++)
                {
                    data += p[i].ToString() + ",";
                }

                data += ";";
            }

            ViewBag.Message += data;

            return View();
        }
    }
}