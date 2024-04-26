using CRM_Raviz.Models;
using OfficeOpenXml;
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

            EventTable eventTable = db.EventTables.Find(int.Parse(form["Id"].ToString()));

            if (true)
            {
                eventTable = new EventTable();
                eventTable.Date = DateTime.Now;
                eventTable.Time = DateTime.Now.TimeOfDay;
                eventTable.CustomerName = form["CustomerName"].ToString();
                eventTable.AccountNo = form["AccountNo"].ToString();
                eventTable.Dispo = form["Disposition"].ToString();
                eventTable.SubDispo = form["SubDisposition"].ToString();
                eventTable.Comments = form["Comments"].ToString();
                eventTable.ChangeStatus = form["ChangeStatus"].ToString();
                if (!string.IsNullOrEmpty(form["CallbackTime"]))
                {
                    eventTable.CallbackTime = DateTime.Parse(form["CallbackTime"]);
                }
                else
                {
                    // Assign a specific default value if CallbackTime is null or empty
                    eventTable.CallbackTime = new DateTime(2000, 1, 1); // Example default value
                }
                eventTable.Record_Id = int.Parse(form["Id"].ToString());

                db.EventTables.Add(eventTable);
            }
            //else
            //{
            //    eventTable.Date = DateTime.Now;
            //    eventTable.Time = BitConverter.GetBytes(DateTime.Now.ToBinary());
            //    eventTable.CustomerName = form["CustomerName"].ToString();
            //    eventTable.AccountNo = form["AccountNo"].ToString();
            //    eventTable.Dispo = form["Disposition"].ToString();
            //    eventTable.SubDispo = form["SubDisposition"].ToString();
            //    eventTable.Comments = form["Comments"].ToString();
            //    eventTable.ChangeStatus = form["ChangeStatus"].ToString();
            //    eventTable.CallbackTime = DateTime.Parse(form["CallbackTime"]);
            //    eventTable.Record_Id = int.Parse(form["Id"].ToString());

            //    db.Entry(eventTable).State = System.Data.Entity.EntityState.Modified;
            //}

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
            if (!string.IsNullOrEmpty(form["CallbackTime"]))
            {
                eventTable.CallbackTime = DateTime.Parse(form["CallbackTime"]);
            }
            else
            {
                // Assign a specific default value if CallbackTime is null or empty
                eventTable.CallbackTime = new DateTime(2000, 1, 1); // Example default value
            }

            db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;



            
            db.SaveChanges();
            return RedirectToAction("EditAllocation");
        }

        [HttpPost]
        public ActionResult _Records(string query)
        {
            CPVDBEntities db = new CPVDBEntities();

            RecordData recordData = new RecordData();

            var results1 = db.RecordDatas.ToList();

            if (query != "")
            {
                results1 = db.RecordDatas
                .Where(item => item.CustomerName == query ||
                               item.AccountNo == query)
                .ToList();

            }
            else
            {
                results1 = db.RecordDatas.ToList();
            }


            return PartialView("_Records", results1);
        }

        public ActionResult _History()
        {

            CPVDBEntities db = new CPVDBEntities();
            List<EventTable> cases = db.EventTables.ToList();
            return PartialView(cases);
        }


        [HttpPost]
        public ActionResult _History(string query)
        {
            CPVDBEntities db = new CPVDBEntities();
            EventTable eventTable = new EventTable();

            var results2 = db.EventTables.ToList();

            if (query != "")
            {

                results2 = db.EventTables
                .Where(item => item.CustomerName == query ||
                               item.AccountNo == query)
                .ToList();

            }
            else
            {
                results2 = db.EventTables.ToList();
            }

            return PartialView("_History", results2);
        }

        [HttpPost]
        public ActionResult UploadCases(HttpPostedFileBase file)
        {
            CPVDBEntities db = new CPVDBEntities();
            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var caseEntity = new RecordData
                            {
                                AccountNo = worksheet.Cells[row, 1].Value?.ToString(),
                                CustomerName = worksheet.Cells[row, 2].Value?.ToString(),
                                BCheque = worksheet.Cells[row, 3].Value?.ToString(),
                                BCheque_P = worksheet.Cells[row, 4].Value?.ToString(),
                                IPTelephone_Billing = worksheet.Cells[row, 5].Value?.ToString(),
                                Utility_Billing = worksheet.Cells[row, 6].Value?.ToString(),
                                Others = worksheet.Cells[row, 7].Value?.ToString(),
                                OS_Billing = worksheet.Cells[row, 8].Value?.ToString(),
                                License_expiry = worksheet.Cells[row, 9].Value?.ToString(),
                                Contact_Person = worksheet.Cells[row, 11].Value?.ToString(),
                                Nationality = worksheet.Cells[row, 12].Value?.ToString(),
                                Mobile1 = worksheet.Cells[row, 13].Value?.ToString(),
                                Mobile2 = worksheet.Cells[row, 14].Value?.ToString(),
                                Mobile3 = worksheet.Cells[row, 15].Value?.ToString(),
                                Mobile4 = worksheet.Cells[row, 16].Value?.ToString(),
                                Email_1 = worksheet.Cells[row, 17].Value?.ToString(),
                                Email_2 = worksheet.Cells[row, 18].Value?.ToString(),
                                Email_3 = worksheet.Cells[row, 19].Value?.ToString(),
                            };

                            db.RecordDatas.Add(caseEntity);
                        }

                        db.SaveChanges();
                    }
                    ViewBag.Message = "File uploaded successfully.";
                }
                catch (Exception ex)
                {
                    ViewBag.Error = $"Error: {ex.Message}";
                }
            }
            else
            {
                ViewBag.Error = "Please upload a valid Excel file.";
            }

            return View();
        }


        public ActionResult UploadCases()
        {

            return View();

        }
    }
}