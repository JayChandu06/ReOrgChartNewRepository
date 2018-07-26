using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace REORGCHART.Controllers
{
    public class ExportController : Controller
    {
        // GET: Export
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult UploadData()
        {
            return View();
        }

    }
}