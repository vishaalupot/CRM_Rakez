using CRM_Raviz.Models;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity.Owin;

namespace CRM_Raviz.Controllers
{
    public class HomeController : Controller
    {
        private ApplicationUserManager _userManager;

        public ActionResult UserMaster()
        {
            CPVDBEntities db = new CPVDBEntities();
            List<string> listRoles = db.AspNetRoles.Select(s => s.Name).ToList();
            ViewBag.Roles = listRoles;
            return View();


        }

        [HttpPost]
        public ActionResult DeleteUser(string id)
        {
            using (CPVDBEntities db = new CPVDBEntities())
            {

                AspNetUser aspNetUser = db.AspNetUsers.Find(id);



                if (aspNetUser != null)
                {
                    db.AspNetUsers.Remove(aspNetUser);
                }
                db.SaveChanges();
                List<AspNetUser> users = db.AspNetUsers.ToList();
                return View(users);
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> UserMaster(RegisterViewModel model)
        {

            var roleManager = new Microsoft.AspNet.Identity.RoleManager<IdentityRole>(new RoleStore<IdentityRole>());
            List<string> roles = roleManager.Roles.Select(r => r.Name.Trim()).ToList();
            ViewBag.Roles = roles;

            if (ModelState.IsValid)
            {
                var user = new ApplicationUser
                {
                    UserName = model.UserName,
                    FullName = model.FullName,
                    UserRole = model.InitialRole,
                };

                model.UserRole = model.InitialRole;
                var result = await UserManager.CreateAsync(user, model.Password);

                if (result.Succeeded)
                {
                    await UserManager.AddToRoleAsync(user.Id, model.UserRole);

                    //await SignInManager.SignInAsync(user, isPersistent: false, rememberBrowser: false);
                    return RedirectToAction("UserMaster", "Home");
                }
                //AddErrors(result);
            }
            return View(model);
        }

        public ApplicationUserManager UserManager
        {
            get
            {
                return _userManager ?? HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            }
            private set
            {
                _userManager = value;
            }
        }


        public ActionResult UserView()
        {
            CPVDBEntities db = new CPVDBEntities();
            List<AspNetUser> users = db.AspNetUsers.ToList();
            return View(users);
        }

        public ActionResult CRMPPF()
        {
            CPVDBEntities db = new CPVDBEntities();
            EventTable eventTable = new EventTable();

            var AgentNames = db.AspNetUsers
                         .Where(r => r.UserRole == "Agent")
                         .Select(r => r.UserName)
                         .Distinct()
                         .ToList();

            // Pass the dispositionValues to the view
            ViewBag.AgentNames = AgentNames;

            return View(eventTable);

        }
        public ActionResult MasterReport()
        {

            CPVDBEntities db = new CPVDBEntities();
            RecordData cases = new RecordData();

            //var dispositionValues = db.RecordDatas.Select(r => r.Disposition).Distinct().ToList();
            //// Pass the dispositionValues to the view
            //ViewBag.DispositionValues = dispositionValues;




            return View(cases);
        }

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
            List<EventTable> cases = db.EventTables.ToList();
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
                eventTable.Agent = form["Agent"].ToString();
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


        public static List<string> GetFieldNames<T>() where T : class
        {
            List<string> fieldNames = new List<string>();

            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                fieldNames.Add(property.Name);
            }

            return fieldNames;
        }

        public ActionResult DownloadCases(string Users, string callbackLanguage, DateTime? specificDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<EventTable> caseTables;
            if(Users != "")
            {
                caseTables = db.EventTables.Where(et => et.Agent == Users ).ToList();

            }
            else if (specificDate.HasValue)
            {
                caseTables = db.EventTables.Where(et => et.Date == specificDate.Value.Date && et.SubDispo == callbackLanguage).ToList();
            }
            else
            {
                caseTables = db.EventTables.Where(et => et.SubDispo == callbackLanguage).ToList();
            }

            var header = new List<string>() { "Id", "AccountNo", "CustomerName", "Date", "Time", "Dispo", "SubDispo", "Comments", "ChangeStatus", "CallbackTime", "Record_Id" };

            using (var package = new ExcelPackage())
            {
                int col = 1;
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                foreach (var headerName in header)
                {
                    worksheet.Cells[1, col++].Value = headerName.ToString();
                }

                int row = 2;
                foreach (var itemcase in caseTables)
                {
                    col = 1;
                    foreach (var headName in header)
                    {
                        var property = typeof(EventTable).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);
                        object value = property.GetValue(itemcase);

                        if (property.PropertyType == typeof(DateTime))
                        {
                            worksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            worksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                    row++;
                }

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Caselist" + (specificDate.HasValue ? specificDate.Value.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd")) + ".xlsx");

                Response.BinaryWrite(package.GetAsByteArray());
                Response.Flush();
                Response.SuppressContent = true;
                HttpContext.ApplicationInstance.CompleteRequest();

                return File(Response.OutputStream, Response.ContentType);
            }
        }

        public ActionResult DownloadRecordCases(string Disposition, DateTime? specificDate = null, DateTime? endDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            if (specificDate.HasValue && endDate.HasValue)
            {
                recordDatas = db.RecordDatas
                                .Where(et => et.CallbackTime >= specificDate.Value.Date && et.CallbackTime <= endDate.Value.Date && et.Disposition == Disposition)
                                .ToList();
            }
            else if (specificDate.HasValue)
            {
                recordDatas = db.RecordDatas
                                .Where(et => et.CallbackTime == specificDate.Value.Date && et.Disposition == Disposition)
                                .ToList();
            }
            else if(Disposition != null)
            {
                recordDatas = db.RecordDatas
                                .Where(et => et.Disposition == Disposition)
                                .ToList();
            }

            var header = new List<string>() { "AccountNo", "CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing", "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person", "Nationality", "Mobile1", "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "Disposition", "SubDisposition", "Comments", "ChangeStatus", "CallbackTime" };

            using (var package = new ExcelPackage())
            {
                int col = 1;
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                foreach (var headerName in header)
                {
                    worksheet.Cells[1, col++].Value = headerName.ToString();
                }

                int row = 2;
                foreach (var itemcase in recordDatas)
                {
                    col = 1;
                    foreach (var headName in header)
                    {
                        var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);
                        object value = property.GetValue(itemcase);

                        if (property.PropertyType == typeof(DateTime))
                        {
                            worksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            worksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                    row++;
                }

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Caselist" + (specificDate.HasValue ? specificDate.Value.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd")) + ".xlsx");

                Response.BinaryWrite(package.GetAsByteArray());
                Response.Flush();
                Response.SuppressContent = true;
                HttpContext.ApplicationInstance.CompleteRequest();

                return File(Response.OutputStream, Response.ContentType);
            }
        }


        public ActionResult DownloadExcel()
        {
            try
            {
                // Create a new Excel package
                using (var excelPackage = new ExcelPackage())
                {
                    // Add a worksheet
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add column headers
                    string[] headers = {
                            "Account #",
                            "Customer Name",
                            "Bounced cheque",
                            "B.Chq Penalities",
                            "IP Telephone billing",
                            "Utility billing",
                            "Others",
                            "O/S Balance",
                            "License expiry",
                            "Contact Person",
                            "Nationality",
                            "Mobile1",
                            "Mobile2",
                            "Mobile3",
                            "Mobile4",
                            "Mobile5",
                            "Email-1",
                            "Email-2",
                            "Email-3",
                            "Email-4",
                            "Email-5",
                            "Closed Account",
                            "Dormant Account",
                            "Insufficient Funds",
                            "Other Reason",
                            "Signature Irregular",
                            "Technical Reason",
                            "Others"
                        };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = headers[i];
                    }

                    // Save the Excel package to a stream
                    MemoryStream memoryStream = new MemoryStream();
                    excelPackage.SaveAs(memoryStream);

                    // Set the position of the stream back to the beginning
                    memoryStream.Position = 0;

                    // Generate a unique file name based on the current date and time
                    string fileName = "SampleExcel_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";

                    // Return the Excel file as a file download response
                    return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                // For simplicity, we'll just return a generic error message
                return Content("An error occurred while generating the Excel file.");
            }
        }


    }
}