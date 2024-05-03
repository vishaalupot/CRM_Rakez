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
using Microsoft.AspNet.Identity;
using System.Data.Entity.Infrastructure;

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

       



        public ActionResult RealEditAllocation(int id, string AccountNo)
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = db.RecordDatas.Find(id);

            List<BouncedRecord> bouncedRecords = db.BouncedRecords.Where(record => record.AccountNo == AccountNo).ToList();
            ViewBag.Bounce = bouncedRecords;

            return View(recordData);
        }

        [HttpPost]
        public ActionResult RealEditAllocation(FormCollection form)
        {
            CPVDBEntities db = new CPVDBEntities();

            RecordData recordData = db.RecordDatas.Find(int.Parse(form["Id"].ToString()));

            EventTable eventTable = db.EventTables.Find(int.Parse(form["Id"].ToString()));
            string val = form["Disposition"].ToString();

            if (val != "CALLBACK" && val != "CALLBACK LANGUAGE")
            {
                recordData.SubDisposition = " ";
                recordData.CallbackTime = new DateTime(2000, 1, 1);
                eventTable.SubDispo = " ";
                eventTable.CallbackTime = new DateTime(2000, 1, 1);


            }
            else if (val == "CALLBACK LANGUAGE")
            {
                recordData.SubDisposition = form["SubDisposition"].ToString();
                recordData.CallbackTime = DateTime.Now;
                eventTable.SubDispo = form["SubDisposition"].ToString();
                eventTable.CallbackTime = DateTime.Now;

            }
            else if (val == "CALLBACK")
            {
                recordData.SubDisposition = " ";
                recordData.CallbackTime = DateTime.Now;
                eventTable.SubDispo = " ";
                eventTable.CallbackTime = DateTime.Now;

            }



            //string History = @User.Identity.GetUserName() + new string(' ', 25)
            //     + DateTime.Now + new string(' ', 25)
            //     + form["Disposition"].ToString() + " "
            //     + form["SubDisposition"].ToString() +" "
            //     + form["CallbackTime"].ToString() 
            //     +" || Remarks: "+ commentsValue;

            // string Heading = "Username" + new string(' ', 25) + "Date" + new string(' ', 50) + "Disposition";



           
                eventTable.Datetime = DateTime.Now;
                eventTable.CustomerName = form["CustomerName"].ToString();
                eventTable.Agent = form["Agent"].ToString();
                eventTable.AccountNo = form["AccountNo"].ToString();
                eventTable.Dispo = form["Disposition"].ToString();
                eventTable.Comments = form["CommentsBox"].ToString();
                //if (!string.IsNullOrEmpty(form["CallbackTime"]))
                //{
                //    eventTable.CallbackTime = DateTime.Parse(form["CallbackTime"]);
                //}
                //else
                //{
                //    eventTable.CallbackTime = new DateTime(2000, 1, 1); // Example default value
                //}
                //eventTable.ChangeStatus = form["ChangeStatus"].ToString();
                //eventTable.ModifiedDate = DateTime.Now;
                eventTable.Segments = recordData.Segments;
                eventTable.Record_Id = int.Parse(form["Id"].ToString());

                db.EventTables.Add(eventTable);


                recordData.Disposition = form["Disposition"].ToString();
                recordData.Comments = form["CommentsBox"].ToString();
                recordData.ChangeStatus = form["ChangeStatus"].ToString();
                recordData.ModifiedDate = DateTime.Now;
                if (!string.IsNullOrEmpty(form["CallbackTime"]))
                {
                recordData.CallbackTime = DateTime.Parse(form["CallbackTime"]);
                }
                else
                {
                // Assign a specific default value if CallbackTime is null or empty
                recordData.CallbackTime = new DateTime(2000, 1, 1); // Example default value
                }

                db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;




            db.SaveChanges();
            return RedirectToAction("RealEditAllocation");
        }


        //public ActionResult EditInfo()
        //{
        //    CPVDBEntities db = new CPVDBEntities();
        //    var results1 = db.RecordDatas.ToList();
        //    results1 = db.RecordDatas
        //           .Where(item => item.Agent == @User.Identity.GetUserName())
        //           .ToList();
        //    return PartialView(results1);
        //}

        //[HttpPost]
        //public ActionResult EditInfo(string value, string modelProperty)
        //{
        //    // Get the model instance from the repository or service
        //    CPVDBEntities db = new CPVDBEntities();
        //    RecordData recordData = new RecordData();

        //    // Update the corresponding property of the model
        //    switch (modelProperty)
        //    {
        //        case "Mobile1":
        //            recordData.Mobile1 = value;
        //            break;
        //        case "Mobile2":
        //            recordData.Mobile2 = value;
        //            break;
        //         case "Mobile3":
        //            recordData.Mobile3 = value;
        //            break;
        //        case "Mobile4":
        //            recordData.Mobile4 = value;
        //            break;
        //        case "Email_1":
        //            recordData.Email_1 = value;
        //            break;
        //        case "Email_2":
        //            recordData.Email_2 = value;
        //            break;
        //        case "Email_3":
        //            recordData.Email_3 = value;
        //            break;
                
        //        default:
        //            break;
        //    }


        //    db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;
        //    db.SaveChanges();
        //    // Return a success response or perform any other necessary actions
        //    return Json(new { success = true });
        //}



        [HttpPost]
        public ActionResult EditInfo(string value, string modelProperty, int id)
        {
            using (var db = new CPVDBEntities())
            {
                var recordData = db.RecordDatas.Find(id); // Load the record from the database

                if (recordData == null)
                {
                    return HttpNotFound(); // Handle case where record doesn't exist
                }

                // Update the corresponding property of the model
                switch (modelProperty)
                {
                    case "Mobile1":
                        recordData.Mobile1 = value;
                        break;
                    case "Mobile2":
                        recordData.Mobile2 = value;
                        break;
                    case "Mobile3":
                        recordData.Mobile3 = value;
                        break;
                    case "Mobile4":
                        recordData.Mobile4 = value;
                        break;
                    case "Email_1":
                        recordData.Email_1 = value;
                        break;
                    case "Email_2":
                        recordData.Email_2 = value;
                        break;
                    case "Email_3":
                        recordData.Email_3 = value;
                        break;

                    default:
                        break;
                }

                try
                {
                    db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                    return Json(new { success = true });
                }
                catch (DbUpdateConcurrencyException ex)
                {
                    // Handle concurrency conflict
                    foreach (var entry in ex.Entries)
                    {
                        entry.Reload(); // Reload the entity from the database
                    }
                    return Json(new { success = false, message = "Concurrency conflict occurred. Please try again." });
                }
            }
        }

        public ActionResult _Records()
        {
            CPVDBEntities db = new CPVDBEntities();
            var results1 = db.RecordDatas.ToList();
            if (User.IsInRole("Agent"))
            {
                results1 = db.RecordDatas
                   .Where(item => item.Agent == @User.Identity.GetUserName())
                   .ToList();
            }
            return PartialView(results1);
        }

        [HttpPost]
        public ActionResult _Records(string query)
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = new RecordData();
            var results1 = db.RecordDatas.ToList();
            var userName = User.Identity.GetUserName();

            if (query == "")
            {
                if (User.IsInRole("Agent"))
                {
                    results1 = db.RecordDatas
                       .Where(item => item.Agent == userName)
                       .ToList();
                }

            }
            else if (query != "")
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

        //[HttpPost]
        //public ActionResult UploadCases(HttpPostedFileBase file)
        //{
        //    CPVDBEntities db = new CPVDBEntities();
        //    if (file != null && file.ContentLength > 0)
        //    {
        //        try
        //        {
        //            using (var package = new ExcelPackage(file.InputStream))
        //            {
        //                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        //                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        //                int rowCount = worksheet.Dimension.Rows;

        //                for (int row = 1; row <= rowCount; row++)
        //                {
        //                    var caseEntity = new RecordData
        //                    {
        //                        AccountNo = worksheet.Cells[row, 1].Value?.ToString(),
        //                        CustomerName = worksheet.Cells[row, 2].Value?.ToString(),
        //                        BCheque = worksheet.Cells[row, 3].Value?.ToString(),
        //                        BCheque_P = worksheet.Cells[row, 4].Value?.ToString(),
        //                        IPTelephone_Billing = worksheet.Cells[row, 5].Value?.ToString(),
        //                        Utility_Billing = worksheet.Cells[row, 6].Value?.ToString(),
        //                        Others = worksheet.Cells[row, 7].Value?.ToString(),
        //                        OS_Billing = worksheet.Cells[row, 8].Value?.ToString(),
        //                        License_expiry = worksheet.Cells[row, 9].Value?.ToString(),
        //                        Contact_Person = worksheet.Cells[row, 11].Value?.ToString(),
        //                        Nationality = worksheet.Cells[row, 12].Value?.ToString(),
        //                        Mobile1 = worksheet.Cells[row, 13].Value?.ToString(),
        //                        Mobile2 = worksheet.Cells[row, 14].Value?.ToString(),
        //                        Mobile3 = worksheet.Cells[row, 15].Value?.ToString(),
        //                        Mobile4 = worksheet.Cells[row, 16].Value?.ToString(),
        //                        Email_1 = worksheet.Cells[row, 17].Value?.ToString(),
        //                        Email_2 = worksheet.Cells[row, 18].Value?.ToString(),
        //                        Email_3 = worksheet.Cells[row, 19].Value?.ToString(),
        //                    };

        //                    db.RecordDatas.Add(caseEntity);
        //                }

        //                db.SaveChanges();
        //            }
        //            ViewBag.Message = "File uploaded successfully.";
        //        }
        //        catch (Exception ex)
        //        {
        //            ViewBag.Error = $"Error: {ex.Message}";
        //        }
        //    }
        //    else
        //    {
        //        ViewBag.Error = "Please upload a valid Excel file.";
        //    }

        //    return View();
        //}


        [HttpPost]
        public ActionResult UploadCases(HttpPostedFileBase file, string dropdown)
        {
            CPVDBEntities dbSheet1 = new CPVDBEntities(); // Database context for the first sheet
            CPVDBEntities dbSheet2 = new CPVDBEntities(); // Database context for the second sheet

            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                        // Process the first sheet
                        ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Overall Allocation"];
                        int rowCount1 = worksheet1.Dimension.Rows;

                        for (int row = 2; row <= rowCount1; row++) // Start from row 2 to skip headers
                        {
                            var caseEntity1 = new RecordData
                            {
                                AccountNo = worksheet1.Cells[row, 1].Value?.ToString(),
                                CustomerName = worksheet1.Cells[row, 2].Value?.ToString(),
                                BCheque = worksheet1.Cells[row, 3].Value?.ToString(),
                                BCheque_P = worksheet1.Cells[row, 4].Value?.ToString(),
                                IPTelephone_Billing = worksheet1.Cells[row, 5].Value?.ToString(),
                                Utility_Billing = worksheet1.Cells[row, 6].Value?.ToString(),
                                Others = worksheet1.Cells[row, 7].Value?.ToString(),
                                OS_Billing = worksheet1.Cells[row, 8].Value?.ToString(),
                                License_expiry = worksheet1.Cells[row, 9].Value?.ToString(),
                                Contact_Person = worksheet1.Cells[row, 10].Value?.ToString(),
                                Nationality = worksheet1.Cells[row, 11].Value?.ToString(),
                                Mobile1 = worksheet1.Cells[row, 12].Value?.ToString(),
                                Mobile2 = worksheet1.Cells[row, 13].Value?.ToString(),
                                Mobile3 = worksheet1.Cells[row, 14].Value?.ToString(),
                                Mobile4 = worksheet1.Cells[row, 15].Value?.ToString(),
                                Email_1 = worksheet1.Cells[row, 16].Value?.ToString(),
                                Email_2 = worksheet1.Cells[row, 17].Value?.ToString(),
                                Email_3 = worksheet1.Cells[row, 18].Value?.ToString(),
                                CloseAccount = worksheet1.Cells[row, 19].Value?.ToString(),
                                DormantAccount = worksheet1.Cells[row, 20].Value?.ToString(),
                                InsufficientFunds = worksheet1.Cells[row, 21].Value?.ToString(),
                                OtherReason= worksheet1.Cells[row, 22].Value?.ToString(),
                                SignatureIrregular = worksheet1.Cells[row, 23].Value?.ToString(),
                                TechnicalReason = worksheet1.Cells[row, 24].Value?.ToString(),
                                BOthers = worksheet1.Cells[row, 25].Value?.ToString(),
                                Agent = worksheet1.Cells[row, 26].Value?.ToString(),
                                Segments = dropdown
                            };

                            dbSheet1.RecordDatas.Add(caseEntity1);
                        }

                        dbSheet1.SaveChanges();

                        // Process the second sheet
                        ExcelWorksheet worksheet2 = package.Workbook.Worksheets["BC Details"];
                        int rowCount2 = worksheet2.Dimension.Rows;

                        for (int row = 2; row <= rowCount2; row++) // Start from row 2 to skip headers
                        {
                            var caseEntity2 = new BouncedRecord // Assuming you have a different model for the second sheet
                            {
                                AccountNo = worksheet2.Cells[row, 1].Value?.ToString(),
                                ChequeNumber = worksheet2.Cells[row, 2].Value?.ToString(),
                                DateBounced = (DateTime?)worksheet2.Cells[row, 3].Value,
                                TotalAmount = worksheet2.Cells[row, 4].Value?.ToString(),
                                ReasonCode = worksheet2.Cells[row, 5].Value?.ToString(),
                                Text = worksheet2.Cells[row, 6].Value?.ToString(),
                                //ChequeDate = (DateTime)worksheet2.Cells[row, 6].Value,

                            };

                            dbSheet2.BouncedRecords.Add(caseEntity2);
                        }



                        dbSheet2.SaveChanges();
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
                caseTables = db.EventTables.Where(et => et.Datetime == specificDate.Value.Date && et.SubDispo == callbackLanguage).ToList();
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