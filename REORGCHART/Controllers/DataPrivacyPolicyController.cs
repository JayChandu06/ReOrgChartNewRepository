using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DashBoardApp.Controllers
{
    public class DataPrivacyPolicyController : Controller
    {
        // GET: DataPrivacyPolicy
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult DataPrivacy()
        {
            return View();
        }
        public ActionResult TermsOfService()
        {
            return View();
        }
        public ActionResult PrivacyPolicy()
        {
            return View();
        }
    }
}