using CRM_Raviz.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CRM_Raviz.Controllers
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

        public ActionResult EditAllocation()
        {
            return View();
        }

        public ActionResult _Records()
        {

            CPVDBEntities db = new CPVDBEntities();
            List<RecordData> cases = db.RecordDatas.ToList();
            return PartialView(cases);
        }

        public ActionResult RealEditAllocation(int id)
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = db.RecordDatas.Find(id);
            return View(recordData);
        }

        [HttpPost]
        public ActionResult RealEditAllocation(FormCollection form)
        {
            CPVDBEntities db = new CPVDBEntities();

            RecordData recordData = db.RecordDatas.Find(int.Parse(form["Id"].ToString()));

            recordData.CustomerName = form["CustomerName"].ToString();
            recordData.OS_Billing = form["OS_Billing"].ToString();
            recordData.License_expiry = form["License_expiry"].ToString();
            recordData.Nationality = form["Nationality"].ToString();
            recordData.Email_1 = form["Email_1"].ToString();
            recordData.Email_2 = form["Email_2"].ToString();
            recordData.Email_3 = form["Email_3"].ToString();
            recordData.Mobile1 = form["Mobile1"].ToString();
            recordData.Mobile2 = form["Mobile2"].ToString();
            recordData.Mobile3 = form["Mobile3"].ToString();
            recordData.Disposition = form["Disposition"].ToString();
            recordData.SubDisposition = form["SubDisposition"].ToString();
            recordData.Comments = form["Comments"].ToString();
            recordData.ChangeStatus = form["ChangeStatus"].ToString();
            recordData.CallbackTime = DateTime.Parse(form["CallbackTime"]);
            db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;

            db.SaveChanges();
            return RedirectToAction("EditAllocation");
        }

    }
}