using CRM_Raviz.Models;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity;
using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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
using ClosedXML.Excel;
using System.Data.Entity;
using Newtonsoft.Json;
using System.Net;
using System.ComponentModel.Design;


namespace CRM_Raviz.Controllers
{
    [Authorize]
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


        public ActionResult BatchMaster()
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = new RecordData();
            var records = db.RecordDatas.ToList();

            List<string> allbatches = records
                   .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                   .Select(r => r.DerbyBatch)
                   .Distinct()
                   .OrderBy(DerbyBatch => int.Parse(DerbyBatch.Split(' ')[1]))
                   .ToList();

            //var batchStatuses = records
            //.Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
            //.GroupBy(r => r.DerbyBatch)
            //.Select(g => new
            //{
            //    Batch = g.Key,
            //    Status = g.Select(r => r.Status).Distinct().SingleOrDefault()
            //})
            //.ToList();

            List<BatchStatus> batchStatuses = records
            .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
            .GroupBy(r => r.DerbyBatch)
            .Select(g => new BatchStatus
            {
                Batch = g.Key,
                Status = g.Select(r => r.Status).Distinct().SingleOrDefault(),
                Count = g.Count()
            })
            .OrderBy(bs => int.Parse(bs.Batch.Split(' ')[1]))
            .ToList();



            List<int> batchWorked = allbatches
            .Select(batch => records
                .Where(r => (!string.IsNullOrEmpty(r.DerbyBatch) && r.ModifiedDate != null && r.DerbyBatch == batch && r.BatchDate == null) || (r.ModifiedDate >= r.BatchDate && r.BatchDate != null && r.DerbyBatch == batch))
                .Count())
            .ToList();

            List<int> totalBatches = allbatches
               .Select(DerbyBatch => records
                   .Count(r => r.DerbyBatch == DerbyBatch))
               .ToList();

            List<int> batchNotWorked = allbatches
            .Select(batch => records
                .Where(r => (!string.IsNullOrEmpty(r.DerbyBatch) && r.ModifiedDate == null && r.DerbyBatch == batch) || (r.ModifiedDate < r.BatchDate && r.BatchDate != null && r.DerbyBatch == batch))
                .Count())
                .ToList();

            ViewBag.batchStatuses = batchStatuses;
            ViewBag.allbatches = allbatches;
            ViewBag.totalBatches = totalBatches;

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
                    return RedirectToAction("UserMaster", "Home");
                }
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
            RecordData recordData = new RecordData();

            var AgentNames = db.AspNetUsers
                         .Where(r => r.UserRole == "Agent")
                         .Select(r => r.UserName)
                         .Distinct()
                         .ToList();
            ViewBag.AgentNames = AgentNames;

            List<string> allbatches = db.RecordDatas
                  .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                  .Select(r => r.DerbyBatch)
                  .Distinct()
                  .OrderBy(DerbyBatch => DerbyBatch)
                  .ToList();

            ViewBag.allbatches = allbatches;

            return View(recordData);

        }
        public ActionResult MasterReport()
        {

            CPVDBEntities db = new CPVDBEntities();
            RecordData cases = new RecordData();

            var AgentNames = db.AspNetUsers
                        .Where(r => r.UserRole == "Agent")
                        .Select(r => r.UserName)
                        .Distinct()
                        .ToList();

            List<string> allbatches = db.RecordDatas
                  .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                  .Select(r => r.DerbyBatch)
                  .Distinct()
                  .OrderBy(DerbyBatch => DerbyBatch)
                  .ToList();

            ViewBag.allbatches = allbatches;
            ViewBag.AgentNames = AgentNames;

            return View(cases);
        }

        public ActionResult ChangeMobile(int id, string newValue)
        {

            CPVDBEntities db = new CPVDBEntities();
            MobileNo mob = db.MobileNos.Find(id);

            mob.Numbers = newValue;


            db.Entry(mob).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            return Json(new { success = true }); 
        }

        public ActionResult AddMobile(int id, string mobileNumber)
        {
            using (CPVDBEntities db = new CPVDBEntities())
            {
                MobileNo newMobile = new MobileNo
                {
                    RecordId = id,  
                    Numbers = mobileNumber
                };

                db.MobileNos.Add(newMobile);
                db.SaveChanges();
            }
            return Json(new { success = true });
        }


        public ActionResult ChangeEmail (int id, string newValue)
        {

            CPVDBEntities db = new CPVDBEntities();
            EmailId mail = db.EmailIds.Find(id);

            mail.Emails = newValue;


            db.Entry(mail).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();

            // Return any response if needed
            return Json(new { success = true }); 
        }

        public ActionResult ChangeStatus(string batch, string status)
        {

            try
            {
                using (CPVDBEntities db = new CPVDBEntities())
                {
                    List<RecordData> records = db.RecordDatas.Where(r => r.DerbyBatch == batch).ToList();

                    foreach (var record in records)
                    {
                        record.Status = status;
                        db.Entry(record).State = System.Data.Entity.EntityState.Modified;
                    }

                    db.SaveChanges();
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                // Log the exception
                // return a JSON response with success set to false and include the exception message
                return Json(new { success = false, message = ex.Message });
            }

            return Json(new { success = true });
        }


        //public ActionResult ChangeStatus(string batch)
        //{

        //    CPVDBEntities db = new CPVDBEntities();
        //    List<RecordData> records = db.RecordDatas.Where(r => r.DerbyBatch == batch).ToList();
        //    //EmailId mail = db.EmailIds.Find(id);

        //    foreach (var record in records)
        //    {
        //        record.Status = "True";
        //    }


        //    db.Entry(records).State = System.Data.Entity.EntityState.Modified;
        //    db.SaveChanges();

        //    // Return any response if needed
        //    return Json(new { success = true });
        //}

        public ActionResult AddEmail(int id, string emailId)
        {
            using (CPVDBEntities db = new CPVDBEntities())
            {
                EmailId email = new EmailId
                {
                    RecordId = id,  
                    Emails = emailId
                };

                db.EmailIds.Add(email);
                db.SaveChanges();
            }
            return Json(new { success = true });
        }

        public ActionResult Index(string filter)
        {
            DateTime today = DateTime.Today;

            using (CPVDBEntities db = new CPVDBEntities())
            {
                var startOfDay = today.Date;
                var endOfDay = today.Date.AddDays(1).AddTicks(-1);
                var startOfMonth = new DateTime(today.Year, today.Month, 1);
                var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);
                var records = db.RecordDatas.Where(r => r.Status == "true").ToList();
                string currentUser = User.Identity.GetUserName();

                IQueryable<RecordData> recordDatasQuery = db.RecordDatas;
                IQueryable<RecordData> ExclusiveRecord = db.RecordDatas;
                IQueryable<EventTable> eventTableQuery = db.EventTables;

                recordDatasQuery = db.RecordDatas.Where(r => r.Status == "true");
                ExclusiveRecord = db.RecordDatas.Where(r => (r.Status) == "true");

                var allAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent")
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();

                var fliterAllAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent")
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();
                ViewBag.Filter = "All";


                if (User.IsInRole("Agent"))
                {
                    recordDatasQuery = db.RecordDatas.Where(r => (r.ModifiedAgent ?? r.Agent) == currentUser);
                    ExclusiveRecord = db.RecordDatas.Where(r => (r.Agent) == currentUser);
                    eventTableQuery = db.EventTables.Where(r => r.Agent == currentUser);

                    allAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent" && r.UserName == currentUser)
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();

                    ViewBag.Filter = currentUser;
                    fliterAllAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent" && r.UserName == currentUser)
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();
                }
                else if (filter == "All")
                {
                    recordDatasQuery = db.RecordDatas;
                    eventTableQuery = db.EventTables;

                    allAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent")
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();

                    ViewBag.Filter = filter;
                }
                else if (filter != "all" && !User.IsInRole("Agent") && filter != null)
                {
                    eventTableQuery = db.EventTables.Where(r => r.Agent == filter);

                    recordDatasQuery = db.RecordDatas.Where(r => (r.ModifiedAgent ?? r.Agent) == filter);
                    ExclusiveRecord = db.RecordDatas.Where(r => (r.Agent) == filter);


                    allAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent" && r.UserName == filter)
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();

                    ViewBag.Filter = filter;
                }
              
                //Total Cases for each Agents
                var totalCasesList = allAgents.Select(UserName => ExclusiveRecord.Count(r => r.Agent == UserName && r.Status =="True" )).ToList(); /*recordDatasQuery.Where().Count();*/

                //Events
                int casesCountToday = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfDay && r.Datetime <= endOfDay)
                     .Select(r => r.AccountNo) 
                    .Distinct()
                    .Count();

                int casesCountThisMonth = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfMonth && r.Datetime <= endOfMonth)
                    .Select(r => r.AccountNo) 
                    .Distinct() 
                    .Count();

                var agentCasesCountToday = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfDay && r.Datetime <= endOfDay)
                    .GroupBy(r => r.Agent)
                    .Select(g => new {
                        Agent = g.Key,
                        CasesCountToday = g.Select(r => r.AccountNo).Distinct().Count()
                        //CasesCountToday = g.Count()
                    })
                    .ToList();

                var agentCasesCountThisMonth = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfMonth && r.Datetime <= endOfMonth)
                    .GroupBy(r => r.Agent)
                    .Select(g => new {
                        Agent = g.Key,
                        CasesCountThisMonth = g.Select(r => r.AccountNo).Distinct().Count()
                    })
                    .ToList();

                var agentsWithCasesCount = allAgents
                    .Select(ModifiedAgent => new AgentCasesViewModel
                    {
                        ModifiedAgent = ModifiedAgent,
                        CasesCountToday = agentCasesCountToday.FirstOrDefault(ac => ac.Agent == ModifiedAgent)?.CasesCountToday ?? 0,
                        CasesCountThisMonth = agentCasesCountThisMonth.FirstOrDefault(ac => ac.Agent == ModifiedAgent)?.CasesCountThisMonth ?? 0
                    })
                    .ToList();

                //Batches+Seg

                List<string> allbatchesComb = records
                   .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                   .Select(r => r.DerbyBatch)
                   .Distinct()
                   .OrderBy(DerbyBatch => int.Parse(DerbyBatch.Split(' ')[1]))
                   .ToList();

                List<int> BCBatches = allbatchesComb
                .Select(batch => ExclusiveRecord
                    .Where(r => !string.IsNullOrEmpty(r.Segments) && r.Segments == "Bounced Cheque" && r.DerbyBatch==batch)
                    .Count())
                .ToList();

                List<int> BBatches = allbatchesComb
                .Select(batch => ExclusiveRecord
                    .Where(r => (!string.IsNullOrEmpty(r.Segments) && r.Segments == "Bounced Cheque and Renewal" && r.DerbyBatch == batch))
                    .Count())
                .ToList();

                List<int> RBatches = allbatchesComb
                .Select(batch => ExclusiveRecord
                    .Where(r => (!string.IsNullOrEmpty(r.Segments) && r.Segments == "Renewal" && r.DerbyBatch == batch))
                    .Count())
                .ToList();

                List<int> totalBatchesComb = allbatchesComb
                   .Select(DerbyBatch => ExclusiveRecord
                       .Count(r => r.DerbyBatch == DerbyBatch))
                   .ToList();

                

                //Batches

                List<string> allbatches = records
                   .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                   .Select(r => r.DerbyBatch)
                   .Distinct()
                   .OrderBy(DerbyBatch => int.Parse(DerbyBatch.Split(' ')[1]))
                   .ToList();

                List<int> batchWorked = allbatches
                .Select(batch => ExclusiveRecord
                    .Where(r => (!string.IsNullOrEmpty(r.DerbyBatch) && r.ModifiedDate >= r.BatchDate && r.DerbyBatch == batch))
                    .Count())
                .ToList();

                List<int> totalBatches = allbatches
                   .Select(DerbyBatch => ExclusiveRecord
                       .Count(r => r.DerbyBatch == DerbyBatch))
                   .ToList();

                List<int> batchNotWorked = allbatches   
                .Select(batch => ExclusiveRecord
                    .Where(r => (!string.IsNullOrEmpty(r.DerbyBatch) && (r.ModifiedDate<r.BatchDate && r.DerbyBatch == batch) || (r.ModifiedDate == null && r.DerbyBatch == batch)))
                    .Count())
                .ToList();

                List<string> allSegmentsComb = records
                .Where(r => !string.IsNullOrEmpty(r.Segments)) // Ensure we only consider non-null and non-empty segments
                .Select(r => r.Segments)
                .Distinct()
                .OrderBy(segment => segment)
                .ToList();

                List<int> segmentCountsComb = allSegmentsComb
                    .Select(segment => ExclusiveRecord
                        .Count(r => r.Segments == segment && r.Status=="True"))
                    .ToList();
               


                //Callback

                int callbackCountThisMonth = recordDatasQuery
                    .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
                    .Count();

                var agentCallBackCountPrev = recordDatasQuery
                    .Where(r => r.CallbackTime.HasValue && r.CallbackTime <= startOfDay && r.CallbackTime != new DateTime(2000, 1, 1))
                    .GroupBy(r => r.ModifiedAgent)
                    .Select(g => new { ModifiedAgent = g.Key, CallBackCountPrev = g.Count() })
                    .ToList();

                var agentCallBackCountToday = recordDatasQuery
                   .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
                   .GroupBy(r => r.ModifiedAgent)
                   .Select(g => new { ModifiedAgent = g.Key, CallBackCountToday = g.Count() })
                   .ToList();

                var agentCallBackCount = allAgents
                    .Select(agent => new AgentCasesViewModel
                    {
                        Agent = agent,
                        CallBackCountToday = agentCallBackCountToday.FirstOrDefault(ac => ac.ModifiedAgent == agent)?.CallBackCountToday ?? 0,
                        CallBackCountPrev = agentCallBackCountPrev.FirstOrDefault(ac => ac.ModifiedAgent == agent)?.CallBackCountPrev ?? 0
                    })
                    .ToList();

                //Segments

                 List<string> allSegments = records
                .Where(r => !string.IsNullOrEmpty(r.Segments)) // Ensure we only consider non-null and non-empty segments
                .Select(r => r.Segments)
                .Distinct()
                .OrderBy(segment => segment)
                .ToList();

                int totalCases = allAgents.Sum(UserName => ExclusiveRecord.Count(r => r.Agent == UserName && r.Status == "True"));

                int totalCasesNotWorked = ExclusiveRecord
                    .Where(r => r.ModifiedDate == null)
                    .Count();
                
                List<int> segmentCounts = allSegments
                    .Select(segment => ExclusiveRecord
                        .Count(r => r.Segments == segment))
                    .ToList();

                List<int> segmentNotWorked = allSegments
                    .Select(segment => ExclusiveRecord
                        .Count(r => r.Segments == segment && r.ModifiedDate == null))
                    .ToList();

                //Pending Cases

                var validDispositions = new[] {
                    "RINGING", "SWITCH OFF", "STATEMENT OF ACCOUNT REQUEST",
                    "THIRD PARTY CALLBACK", "THIRD PARTY CONTACT",
                    "THIRD PARTY CTC INFO UPDATE", "REFUSE TO DE-REGISTER",
                    "REFUSE TO RENEW", "PAYMENT MISSING", "INVALID NUMBER",
                    "LINE BUSY", "CUSTOMER HUNG UP", "CUSTOMER OUT OF COUNTRY",
                    "ACCOUNT EXCLUDED"
                };

                var ringingAccounts1 = db.EventTables
                    .GroupBy(r => r.AccountNo)
                    .Where(g => g.Count() == g.Count(r => validDispositions.Contains(r.Dispo)))
                    .Select(g => g.Key)
                    .ToList();

                var RingingAccounts = db.RecordDatas
                    .Where(r => ringingAccounts1.Contains(r.AccountNo)) 
                    .GroupBy(r => r.Agent ?? r.ModifiedAgent)
                    .Select(g => new
                    {
                        AgentName = g.Key,
                        AccountCount = g.Count()
                    })
                    .ToList();

                //var distinctSegments = recordDatasQuery
                //    .Where(r => !string.IsNullOrEmpty(r.Segments))
                //    .Select(r => r.Segments)
                //    .OrderBy(segment => segment)
                //    .Distinct()
                //    .ToList();

                //var segment2Counts = recordDatasQuery
                //    .Where(r => !string.IsNullOrEmpty(r.Segments))
                //    .GroupBy(r => r.DerbyBatch)
                //    .OrderBy(g => g.Key)
                //    .Select(g => g.Count())
                //    .ToList();



                //var distinctSegments2 = recordDatasQuery
                //    .Where(r => !string.IsNullOrEmpty(r.Segments))
                //    .Select(r => r.DerbyBatch)
                //    .OrderBy(DerbyBatch => DerbyBatch)
                //    .Distinct()
                //    .ToList();


                //var casesRinging = eventTableQuery
                //    .GroupBy(r => r.AccountNo)
                //    .Where(g => g.Count(r => r.Dispo == "RINGING" || r.Dispo == "SWITCH OFF") == g.Count())
                //    .Select(g => new { Agent = g.Key, RingingToday = g.Count() });


                ViewBag.allbatchesComb = allbatchesComb;
                ViewBag.BCBatches = BCBatches;
                ViewBag.BBatches = BBatches;
                ViewBag.RBatches = RBatches;
                ViewBag.totalBatchesComb = totalBatchesComb;
                ViewBag.segmentCountsComb = segmentCountsComb;

                ViewBag.totalCasesList = totalCasesList;
                ViewBag.allAgents = fliterAllAgents;
                ViewBag.Agents = allAgents;
                ViewBag.allbatches = allbatches;
                ViewBag.totalCasesNotWorked = totalCasesNotWorked;
                ViewBag.totalCases = totalCases;
                ViewBag.segmentNotWorked = segmentNotWorked;
                ViewBag.TotalBatches = totalBatches;
                ViewBag.BatchNotWorked = batchNotWorked;
                ViewBag.BatchWorked = batchWorked;
                ViewBag.segmentCounts = segmentCounts;
                ViewBag.recentEvents = casesCountToday;
                ViewBag.callbackCountThisMonth = callbackCountThisMonth;
                ViewBag.CasesCountThisMonth = casesCountThisMonth;
                ViewBag.RingingNos = ringingAccounts1.Count;
                ViewBag.RingingAccounts = RingingAccounts;
                ViewBag.AgentCallBackCount = agentCallBackCount;

                //ViewBag.segment2Counts = segment2Counts;
                //ViewBag.distinctSegments2 = distinctSegments2;
                //ViewBag.distinctSegments = distinctSegments;

                return View(agentsWithCasesCount);
            }
        }

        public ActionResult Ringing(string Agent)
        {
            using (CPVDBEntities db = new CPVDBEntities())
            {
                IQueryable<EventTable> events = db.EventTables;
                IQueryable<RecordData> recordDatas = db.RecordDatas;

                var filteredEventTables = events.Where(r => r.Agent == Agent).ToList();

                var validDispositions = new[] {
                    "RINGING", "SWITCH OFF", "STATEMENT OF ACCOUNT REQUEST",
                    "THIRD PARTY CALLBACK", "THIRD PARTY CONTACT",
                    "THIRD PARTY CTC INFO UPDATE", "REFUSE TO DE-REGISTER",
                    "REFUSE TO RENEW", "PAYMENT MISSING", "INVALID NUMBER",
                    "LINE BUSY", "CUSTOMER HUNG UP", "CUSTOMER OUT OF COUNTRY",
                    "ACCOUNT EXCLUDED"
                };

                var ringingAccounts1 = db.EventTables
                    .GroupBy(r => r.AccountNo)
                    .Where(g => g.Count() == g.Count(r => validDispositions.Contains(r.Dispo)))
                    .Select(g => g.Key) // Select the account numbers
                    .ToList();

                int i = 0;

                var RingingAccounts = recordDatas
                 .Where(r => ringingAccounts1.Contains(r.AccountNo) && r.Agent == Agent)
                 .Select(r => new
                 {
                     Index = i + 1,
                     AccountNo = r.AccountNo,
                     AgentName = r.Agent,
                     CustomerName = r.CustomerName,
                     itemId = r.Id
                 })
                 .Distinct()
                 .ToList();

               

                return Json(RingingAccounts, JsonRequestBehavior.AllowGet);
            }
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

        [HttpPost]
        public ActionResult changedAgent(int id, string Agent)
        {
            CPVDBEntities db = new CPVDBEntities();
            var record = db.RecordDatas.FirstOrDefault(item => item.Id == id);
            if (record != null)
            {
                record.Agent = Agent;
            }

            db.Entry(record).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();

            // Return any response if needed
            return Json(new { success = true }); // For example, returning a JSON response
        }





        public ActionResult RealEditAllocation(int id, string AccountNo)
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = db.RecordDatas.Find(id);
            MobileNo mobileNo =  new MobileNo();
            EmailId emailId = new EmailId();

            List<BouncedRecord> bouncedRecords = db.BouncedRecords.Where(record => record.AccountNo == AccountNo).ToList();
            ViewBag.Bounce = bouncedRecords;


            var mobileData = db.MobileNos.
                Where(item => item.RecordId == id && item.Validity != "Invalid")
                .ToList();


            ViewBag.DialedNumber = mobileData;

            var EmailId = db.EmailIds
                .Where(item => item.RecordId == id)
                .ToList();

            ViewBag.EmailUsed = EmailId;


            return View(recordData);
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult RealEditAllocation(FormCollection form)
        {
            CPVDBEntities db = new CPVDBEntities();


            RecordData recordData = db.RecordDatas.Find(int.Parse(form["Id"].ToString()));

            EventTable eventTable = new EventTable();
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
                recordData.CallbackTime = DateTime.TryParse(form["CallbackTime"], out DateTime parsedDateTime) ? parsedDateTime : (DateTime?)null; eventTable.SubDispo = form["SubDisposition"].ToString();
                eventTable.CallbackTime = DateTime.TryParse(form["CallbackTime"], out DateTime tempDate) ? tempDate : (DateTime?)null;

            }
            else if (val == "CALLBACK")
            {
                recordData.SubDisposition = " ";
                recordData.CallbackTime = DateTime.TryParse(form["CallbackTime"], out DateTime parsedDateTime) ? parsedDateTime : (DateTime?)null; eventTable.SubDispo = form["SubDisposition"].ToString();
                eventTable.SubDispo = " ";
                eventTable.CallbackTime = DateTime.TryParse(form["CallbackTime"], out DateTime tempDate) ? tempDate : (DateTime?)null;

            }



                eventTable.Datetime = DateTime.Now;
                eventTable.CustomerName = form["CustomerName"].ToString();
                eventTable.Agent = form["AgentsName"].ToString();
                
                eventTable.AccountNo = form["AccountNo"].ToString();
                
                eventTable.Comments = form["CommentsBox"].ToString();
                eventTable.EmailUsed = form["EmailUsed"].ToString();
                eventTable.CallType = form["CallType"].ToString();
                eventTable.Segments = recordData.Segments;
                eventTable.Record_Id = int.Parse(form["Id"].ToString());

                db.EventTables.Add(eventTable);


                recordData.CallType = form["CallType"].ToString();

                if (form["CallType"].ToString() == "OUTBOUND" || form["CallType"].ToString() == "INBOUND")
                {
                    eventTable.Dispo = form["Disposition"].ToString();
                    recordData.Disposition = form["Disposition"].ToString();
                }
                else if (form["CallType"].ToString() == "EMAIL UPDATE")
                {
                    eventTable.Dispo = form["DispositionSecond"].ToString();
                    recordData.Disposition = form["DispositionSecond"].ToString();
                }
                else if (form["CallType"].ToString() == "ACCOUNT UPDATE")
                {
                    eventTable.Dispo = form["DispositionThird"].ToString();
                    recordData.Disposition = form["DispositionThird"].ToString();
                }

            if (form["Disposition"].ToString() == "INVALID NUMBER")
            {
                string dialedNumber = form["DialedNumber"];
                MobileNo mobileNo = db.MobileNos.FirstOrDefault(m => m.Numbers == dialedNumber);
                mobileNo.Validity = "Invalid";
                db.Entry(mobileNo).State = System.Data.Entity.EntityState.Modified;

            }


                recordData.DialedNumber = form["DialedNumber"].ToString();
                eventTable.DialedNumber = form["DialedNumber"].ToString();


                recordData.Comments = form["CommentsBox"].ToString();
                recordData.ChangeStatus = form["ChangeStatus"].ToString();
                recordData.ModifiedDate = DateTime.Now;
                recordData.ModifiedAgent = form["AgentsName"].ToString();

            db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;




            db.SaveChanges();
            return RedirectToAction("RealEditAllocation", new { id = form["Id"].ToString(), accountno = form["AccountNo"].ToString() });
        }





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

            //List<string> allbatches = db.RecordDatas
            //      .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
            //      .Select(r => r.DerbyBatch)
            //      .Distinct()
            //      .OrderBy(DerbyBatch => DerbyBatch)
            //      .ToList();

            List<string> allbatches = db.RecordDatas
                .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                .Select(r => r.DerbyBatch)
                .Distinct()
                .OrderBy(batch => int.Parse(System.Text.RegularExpressions.Regex.Match(batch, @"\d+$").Value))
                .ToList();


            var allAgents = db.AspNetUsers
                     .Where(r => r.UserRole == "Agent")
                     .Select(r => r.UserName)
                     .Distinct()
                     .ToList();

            ViewBag.allAgents = allAgents;

            ViewBag.allbatches = allbatches;
            return PartialView(results1);
        }

        

        [HttpPost]
        public ActionResult _Records(string query, string drop1, string drop2, string drop3, string drop4, int page = 1, int pageSize = 100)
        {
            CPVDBEntities db = new CPVDBEntities();
            var userName = User.Identity.GetUserName();


            

            var AgentNames = db.AspNetUsers
                         .Where(r => r.UserRole == "Agent")
                         .Select(r => r.UserName)
                         .Distinct()
                         .ToList();

            // Pass the dispositionValues to the view
            ViewBag.AgentNames = AgentNames;

            IQueryable<RecordData> records = db.RecordDatas;



            switch (drop3)
            {

                case "QUOTATION REQUESTED":
                    records = records.Where(et => et.Disposition == "QUOTATION REQUESTED").AsQueryable();
                    break;
                 case "QUOTATION SENT":
                    records = records.Where(et => et.Disposition == "QUOTATION SENT").AsQueryable();
                    break;
                 case "ACCOUNT EXCLUDED":
                    records = records.Where(et => et.Disposition == "ACCOUNT EXCLUDED").AsQueryable();
                    break;
                
                case "CALLBACK LANGUAGE":
                    records = records.Where(et => et.Disposition == "CALLBACK LANGUAGE").AsQueryable();
                    break;
                case "BILL DISPUTE":
                    records = records.Where(et => et.Disposition == "BILL DISPUTE").AsQueryable();
                    break;
                case "CALLBACK":
                    records = records.Where(et => et.Disposition == "CALLBACK").AsQueryable();
                    break;
                case "BILL DISPUTE REFUSE TO PAY":
                    records = records.Where(et => et.Disposition == "BILL DISPUTE REFUSE TO PAY").AsQueryable();
                    break;
                case "CUSTOMER DECEASED":
                    records = records.Where(et => et.Disposition == "CUSTOMER DECEASED").AsQueryable();
                    break;
                case "CUSTOMER HUNG UP":
                    records = records.Where(et => et.Disposition == "CUSTOMER HUNG UP").AsQueryable();
                    break;
                case "CUSTOMER OUT OF COUNTRY":
                    records = records.Where(et => et.Disposition == "CUSTOMER OUT OF COUNTRY").AsQueryable();
                    break;
                case "DE-REGISTRATION":
                    records = records.Where(et => et.Disposition == "DE-REGISTRATION").AsQueryable();
                    break;
                case "DE-REGISTRATION DONE":
                    records = records.Where(et => et.Disposition == "DE-REGISTRATION DONE").AsQueryable();
                    break;
                case "DO NOT CALL":
                    records = records.Where(et => et.Disposition == "DO NOT CALL").AsQueryable();
                    break;
                case "FOLLOW UP":
                    records = records.Where(et => et.Disposition == "FOLLOW UP").AsQueryable();
                    break;
                case "INVALID NUMBER":
                    records = records.Where(et => et.Disposition == "INVALID NUMBER").AsQueryable();
                    break;
                case "LINE BUSY":
                    records = records.Where(et => et.Disposition == "LINE BUSY").AsQueryable();
                    break;
                case "PAYMENT INSTALLMENT APPROVED":
                    records = records.Where(et => et.Disposition == "PAYMENT INSTALLMENT APPROVED").AsQueryable();
                    break;
                case "PAYMENT INSTALLMENT REQUEST":
                    records = records.Where(et => et.Disposition == "PAYMENT INSTALLMENT REQUEST").AsQueryable();
                    break;
                case "PAYMENT MADE":
                    records = records.Where(et => et.Disposition == "PAYMENT MADE").AsQueryable();
                    break;
                case "PAYMENT MISSING":
                    records = records.Where(et => et.Disposition == "PAYMENT MISSING").AsQueryable();
                    break;
                case "PAYMENT REMINDER":
                    records = records.Where(et => et.Disposition == "PAYMENT REMINDER").AsQueryable();
                    break;
                case "PROMISE TO PAY":
                    records = records.Where(et => et.Disposition == "PROMISE TO PAY").AsQueryable();
                    break;
                case "RECALLED":
                    records = records.Where(et => et.Disposition == "RECALLED").AsQueryable();
                    break;
                case "REFUSE TO DE-REGISTER":
                    records = records.Where(et => et.Disposition == "REFUSE TO DE-REGISTER").AsQueryable();
                    break;
                case "REFUSE TO RENEW":
                    records = records.Where(et => et.Disposition == "REFUSE TO RENEW").AsQueryable();
                    break;
                case "RENEWAL DONE":
                    records = records.Where(et => et.Disposition == "RENEWAL DONE").AsQueryable();
                    break;
                case "RENEWAL INQUIRY":
                    records = records.Where(et => et.Disposition == "RENEWAL INQUIRY").AsQueryable();
                    break;
                case "RINGING":
                    records = records.Where(et => et.Disposition == "RINGING").AsQueryable();
                    break;
                case "STATEMENT OF ACCOUNT REQUEST":
                    records = records.Where(et => et.Disposition == "STATEMENT OF ACCOUNT REQUEST").AsQueryable();
                    break;
                case "SWITCH OFF":
                    records = records.Where(et => et.Disposition == "SWITCH OFF").AsQueryable();
                    break;
                case "THIRD PARTY CALLBACK":
                    records = records.Where(et => et.Disposition == "THIRD PARTY CALLBACK").AsQueryable();
                    break;
                case "THIRD PARTY CONTACT":
                    records = records.Where(et => et.Disposition == "THIRD PARTY CONTACT").AsQueryable();
                    break;
                case "THIRD PARTY CTC INFO UPDATE":
                    records = records.Where(et => et.Disposition == "THIRD PARTY CTC INFO UPDATE").AsQueryable();
                    break;
                default:
                    // Handle default case
                    break;
            }





            if (drop1 == "Recent events")
            {
                records = records.OrderByDescending(et => et.ModifiedDate).AsQueryable(); // Order by ModifiedDate for recent events
            }
            // Add more conditions for other values of drop1 if needed
            else if (drop1 == "Old events")
            {
                records = records.OrderBy(et => et.ModifiedDate).AsQueryable(); // Order by ModifiedDate for old events
            }

            List<string> allbatches = records
                  .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                  .Select(r => r.DerbyBatch)
                  .Distinct()
                  .OrderBy(DerbyBatch => DerbyBatch)
                  .ToList();



            if (allbatches.Contains(drop2))
            {
                records = records.Where(et => et.DerbyBatch == drop2).AsQueryable();
            }

            if(drop3 == "Bounced Cheque")
            {
                records = records.Where(et => et.Segments == "Bounced Cheque").AsQueryable();
            }
            else if (drop3 == "Renewal")
            {
                records = records.Where(et => et.Segments == "Renewal").AsQueryable();
            }
            else if (drop3 == "Bounced Cheque and Renewal")
            {
                records = records.Where(et => et.Segments == "Bounced Cheque and Renewal").AsQueryable();
            }

            if (AgentNames.Contains(drop4))
            {
                records = records.Where(et => et.Agent == drop4).AsQueryable();
            }

           
            if (string.IsNullOrEmpty(query))
            {
                if (User.IsInRole("Agent"))
                {
                    records = records.Where(item => item.Agent == userName);
                }
            }
            else
            {
                records = records.Where(item => item.CustomerName == query || item.AccountNo == query || item.Mobile1 == query || item.Mobile2 == query || item.Mobile3 == query || item.Mobile4 == query || item.Agent == query );
            }

            int totalRecords = records.Count();

            var results = records
                .OrderByDescending(item => item.DerbyBatch) // Order by the latest DerbyBatch
                .ThenBy(item => item.Id) // Ensure stable ordering by Id
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            ViewBag.allbatches = allbatches;

            ViewBag.CurrentPage = page;
            ViewBag.TotalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

            return PartialView("_Records", results);
        }


        [HttpPost]
        public ActionResult _AgentCases(string query, string section, string state, string button)
        {
            CPVDBEntities db = new CPVDBEntities();

            List<RecordData> recordDatas = new List<RecordData>();
            List<EventTable> eventTables = new List<EventTable>();

            IQueryable<RecordData> records = db.RecordDatas.Where(r=> r.Status == "True");
            IQueryable<EventTable> events = db.EventTables;

            bool userExists = db.AspNetUsers
                    .Any(r => r.UserRole == "Agent" && r.UserName == query);

            if (userExists)
            {
                records = records.Where(item => item.Agent == query);
            }

            DateTime today = DateTime.Today;
            var startOfDay = today.Date;
            var endOfDay = today.Date.AddDays(1).AddTicks(-1);
            var startOfMonth = new DateTime(today.Year, today.Month, 1);
            var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);

            if (state == "ringing")
            {

                var settings = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                ViewBag.Data = JsonConvert.SerializeObject(events, settings);


                return PartialView("_AgentCases", events);
            }
            else if (state == "segmentCombo")
            {
                
                    records = records.Where(item => item.Segments == section && item.DerbyBatch == button);

                    var settings = new JsonSerializerSettings
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                    };

                    ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                return PartialView("_AgentCases", records);


            }
            else if (state == "callback")
            {
                if (section == "day")
                {
                    records = records.Where(item => item.ModifiedAgent == query && item.CallbackTime >= startOfDay && item.CallbackTime <= endOfDay);

                    var settings = new JsonSerializerSettings
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                    };

                    ViewBag.Data = JsonConvert.SerializeObject(records, settings);



                }
                else if (section == "prev")
                {
                    records = records.Where(item => item.ModifiedAgent == query && item.CallbackTime < startOfDay && item.CallbackTime != new DateTime(2000, 1, 1));
                    var settings = new JsonSerializerSettings
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                    };

                    ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                }
                return PartialView("_AgentCases", records);
            }
            else if (state == "event")
            {
                if (section == "day")
                {
                    //records = records.Where(item => item.Agent == query && item.ModifiedDate >= startOfDay && item.ModifiedDate <= endOfDay);
                    var filteredRecords = from record in db.RecordDatas
                                          where (from eventRecord in db.EventTables
                                                 where eventRecord.Agent == query &&
                                                                   eventRecord.Datetime >= startOfDay && eventRecord.Datetime <= endOfDay
                                                 select eventRecord.AccountNo).Distinct().Contains(record.AccountNo)
                                          select record;

                    var settings = new JsonSerializerSettings
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                    };

                    ViewBag.Data = JsonConvert.SerializeObject(filteredRecords, settings);


                    return PartialView("_AgentCases", filteredRecords);

                }
                else if (section == "month")
                {
                    var filteredRecords = from record in db.RecordDatas
                                          where (from eventRecord in db.EventTables
                                                 where eventRecord.Agent == query &&
                                                                   eventRecord.Datetime >= startOfMonth && eventRecord.Datetime <= endOfMonth
                                                 select eventRecord.AccountNo).Distinct().Contains(record.AccountNo)
                                          select record;

                    var settings = new JsonSerializerSettings
                    {
                        ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                    };

                    ViewBag.Data = JsonConvert.SerializeObject(filteredRecords, settings);

                    return PartialView("_AgentCases", filteredRecords);
                }

                return PartialView("_AgentCases", records);
            }
            else if (state == "segment1")
            {

                if (section == "all")
                {
                    if (button == "All")
                    {
                        records = records.Where(item => item.Segments == query);
                    }
                    else
                    {
                        records = records.Where(item => item.Segments == query && item.Agent == button);

                    }
                }
                else if (section == "notworked")
                {
                    if (button == "All")
                    {
                        records = records.Where(item => item.Segments == query && item.ModifiedDate == null);
                    }
                    else
                    {
                        records = records.Where(item => item.Segments == query && item.ModifiedDate == null && item.Agent == button);
                    }
                }
                //ViewBag.Data = records;
                var settings = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                return PartialView("_AgentCases", records);

            }
            else if(state == "segment2")
            {
                string currentUser = User.Identity.GetUserName();

                bool isAgent = db.AspNetUsers
                    .Any(r => r.UserRole == "Agent" && r.UserName == currentUser || r.UserRole == "Agent" && r.UserName == button);
                if (isAgent)
                {
                    records = records.Where(item => (item.Agent) == currentUser || ( item.Agent) == button);
                }

                if (section == "notworked")
                {
                    records = records.Where(item => (item.DerbyBatch == query && item.ModifiedDate == null) || (item.ModifiedDate < item.BatchDate && item.BatchDate != null && item.DerbyBatch == query));
                }
                else if (section == "worked")
                {
                    records = records.Where(item => (item.ModifiedDate >= item.BatchDate  && item.DerbyBatch == query));
                }
                var settings = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                return PartialView("_AgentCases", records);
            }
            else
            {
                records = records.Where(item => item.Agent == query);
                var settings = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                return PartialView("_AgentCases", records);
            }

        }

       


        [HttpPost]
        public ActionResult _ProfilePic(HttpPostedFileBase file, string snr)
        {


            CPVDBEntities db = new CPVDBEntities();

            var userRole = db.AspNetUsers.Find(User.Identity.GetUserId());
            ViewBag.userRole = userRole.UserRole;

            if (file == null)
            {
                ModelState.AddModelError("", "Please attached a file.");
            }

            if (ModelState.IsValid)
            {

               

                List<AspNetUser> user = db.AspNetUsers.ToList();

                AspNetUser aspNetUser = db.AspNetUsers.Find(snr.ToString());


                if (aspNetUser != null)
                {
                    aspNetUser.Images = ConvertToBytes(file);
                    db.Entry(aspNetUser).State = EntityState.Modified;
                    db.SaveChanges();
                }
                return PartialView();
            }
            else
            {
                return PartialView();
            }

        }

        public byte[] ConvertToBytes(HttpPostedFileBase image)
        {
            byte[] imageBytes = null;
            BinaryReader reader = new BinaryReader(image.InputStream);
            imageBytes = reader.ReadBytes((int)image.ContentLength);
            return imageBytes;
        }

        public ActionResult _ProfilePic()
        {
            CPVDBEntities db = new CPVDBEntities();
            var userRole = db.AspNetUsers.Find(User.Identity.GetUserId());
            ViewBag.userRole = userRole.UserRole;
            return PartialView();
        }

        [HttpGet]
        public ActionResult DisplayProfilePic(string snr)
        {
            CPVDBEntities db = new CPVDBEntities();
            AspNetUser user = db.AspNetUsers.Find(snr);

            if (user != null && user.Images != null)
            {
                return File(user.Images, "image/jpeg"); // Adjust content type based on your image type
            }
            else
            {
                // Return a default image or handle the case when the image is not found
                // For now, return an empty image
                byte[] emptyImage = new byte[0];
                return File(emptyImage, "image/jpeg");
            }
        }

        [ValidateInput(false)]
        public ActionResult DownloadAgentCases(string records)
        {
            try
            {
                string unescapedJson = JsonConvert.DeserializeObject<string>(records);

                // Then, deserialize the unescaped JSON string into a list of objects
                List<RecordData> recordList = JsonConvert.DeserializeObject<List<RecordData>>(unescapedJson);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Agent Cases");
                    var currentRow = 1;

                    // Check if the list has elements
                    if (recordList.Any())
                    {
                        // Get the type of the first element
                        var recordType = recordList.First().GetType();

                        // Get properties of the type
                        var properties = recordType.GetProperties();

                        // Add column headers
                        for (int i = 0; i < properties.Length; i++)
                        {
                            if(properties[i].Name != "EmailIds" && properties[i].Name != "MobileNos")
                            {
                                worksheet.Cells[currentRow, i + 1].Value = properties[i].Name;

                            }
                        }

                        // Add data rows
                        foreach (var record in recordList)
                        {
                            currentRow++;
                            for (int i = 0; i < properties.Length; i++)
                            {
                                if (properties[i].Name != "EmailIds" && properties[i].Name != "MobileNos")
                                {
                                    worksheet.Cells[currentRow, i + 1].Value = properties[i].GetValue(record)?.ToString();

                                }
                                }
                        }
                    }

                    using (var stream = new MemoryStream())
                    {
                        package.SaveAs(stream);
                        stream.Position = 0;

                        var fileName = $"AgentCases_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                        Response.Clear();
                        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        Response.AddHeader("content-disposition", "attachment;filename= " + fileName);
                        Response.BinaryWrite(package.GetAsByteArray());
                        Response.Flush();
                        Response.SuppressContent = true;
                        HttpContext.ApplicationInstance.CompleteRequest();
                        string ipAddress = System.Web.HttpContext.Current.Request.UserHostAddress;
                        return File(Response.OutputStream, Response.ContentType);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                return Content($"An error occurred: {ex.Message}");
            }
        }


        public ActionResult _History()
        {

            CPVDBEntities db = new CPVDBEntities();
            List<EventTable> cases = db.EventTables.OrderByDescending(item => item.Id) // Assuming EventDate is the property representing the time of the events
                 .ToList();
            return PartialView(cases);
        }


        [HttpPost]
        public ActionResult _History(string query)
        {
            CPVDBEntities db = new CPVDBEntities();
            EventTable eventTable = new EventTable();

            var results2 = db.EventTables.OrderByDescending(item => item.Id) // Assuming EventDate is the property representing the time of the events
                 .ToList();

            if (query != "")
            {

                 results2 = db.EventTables
                 .Where(item => item.CustomerName == query || item.AccountNo == query)
                 .OrderByDescending(item => item.Id) // Assuming EventDate is the property representing the time of the events
                 .ToList();


            }
            else
            {
                results2 = db.EventTables
                    .Where(item => item.AccountNo == query)
                    .ToList();
            }

            return PartialView("_History", results2);
        }
        [HttpPost]
        [Authorize]
        public ActionResult UploadCases(HttpPostedFileBase file, string dropdown)
        {
            CPVDBEntities dbSheet1 = new CPVDBEntities(); // Database context for the first sheet
            CPVDBEntities dbSheet2 = new CPVDBEntities(); // Database context for the second sheet
            CPVDBEntities db = new CPVDBEntities(); // Database context for the second sheet

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
                            var accountNo = worksheet1.Cells[row, 1].Value?.ToString();

                            // Check if the AccountNo exists in the RecordData table
                            bool accountExists = dbSheet1.RecordDatas.Any(rd => rd.AccountNo == accountNo);
                            if (accountExists)
                            {
                                var existingRecord = dbSheet1.RecordDatas.SingleOrDefault(rd => rd.AccountNo == accountNo);
                                if (existingRecord != null)
                                {
                                    existingRecord.BatchHistory = existingRecord.BatchHistory + existingRecord.DerbyBatch;
                                    existingRecord.CustomerName = worksheet1.Cells[row, 2].Value?.ToString();
                                    existingRecord.Contact_Person = worksheet1.Cells[row, 3].Value?.ToString();
                                    existingRecord.Nationality = worksheet1.Cells[row, 4].Value?.ToString();
                                    existingRecord.Mobile1 = worksheet1.Cells[row, 5].Value?.ToString();
                                    existingRecord.Mobile2 = worksheet1.Cells[row, 6].Value?.ToString();
                                    existingRecord.Mobile3 = worksheet1.Cells[row, 7].Value?.ToString();
                                    existingRecord.Mobile4 = worksheet1.Cells[row, 8].Value?.ToString();
                                    existingRecord.Email_1 = worksheet1.Cells[row, 9].Value?.ToString();
                                    existingRecord.Email_2 = worksheet1.Cells[row, 10].Value?.ToString();
                                    existingRecord.Email_3 = worksheet1.Cells[row, 11].Value?.ToString();
                                    existingRecord.TenacyFacilityType = worksheet1.Cells[row, 12].Value?.ToString();
                                    existingRecord.License_expiry = worksheet1.Cells[row, 13].Value?.ToString();
                                    existingRecord.ExpectedRenewalFee = worksheet1.Cells[row, 14].Value?.ToString();
                                    existingRecord.SRNumber = worksheet1.Cells[row, 15].Value?.ToString();
                                    existingRecord.DeRegFee = worksheet1.Cells[row, 16].Value?.ToString();
                                    existingRecord.BCheque = worksheet1.Cells[row, 17].Value?.ToString();
                                    existingRecord.BCheque_P = worksheet1.Cells[row, 18].Value?.ToString();
                                    existingRecord.IPTelephone_Billing = worksheet1.Cells[row, 19].Value?.ToString();
                                    existingRecord.Utility_Billing = worksheet1.Cells[row, 20].Value?.ToString();
                                    existingRecord.Others = worksheet1.Cells[row, 21].Value?.ToString();
                                    existingRecord.OS_Billing = worksheet1.Cells[row, 22].Value?.ToString();
                                    existingRecord.CloseAccount = worksheet1.Cells[row, 23].Value?.ToString();
                                    existingRecord.DormantAccount = worksheet1.Cells[row, 24].Value?.ToString();
                                    existingRecord.InsufficientFunds = worksheet1.Cells[row, 25].Value?.ToString();
                                    existingRecord.OtherReason = worksheet1.Cells[row, 26].Value?.ToString();
                                    existingRecord.SignatureIrregular = worksheet1.Cells[row, 27].Value?.ToString();
                                    existingRecord.TechnicalReason = worksheet1.Cells[row, 28].Value?.ToString();
                                    existingRecord.BOthers = worksheet1.Cells[row, 29].Value?.ToString();
                                    existingRecord.EmployeeVisaQuota = worksheet1.Cells[row, 30].Value?.ToString();
                                    existingRecord.EmployeeVisaUtilized = worksheet1.Cells[row, 31].Value?.ToString();
                                    existingRecord.ProjectBundleName = worksheet1.Cells[row, 32].Value?.ToString();
                                    existingRecord.LicenseType = worksheet1.Cells[row, 33].Value?.ToString();
                                    existingRecord.FacilityType = worksheet1.Cells[row, 34].Value?.ToString();
                                    existingRecord.NoYears = worksheet1.Cells[row, 35].Value?.ToString();
                                    existingRecord.DerbyBatch = worksheet1.Cells[row, 36].Value?.ToString();
                                    existingRecord.Agent = worksheet1.Cells[row, 37].Value?.ToString();
                                    existingRecord.BatchDate = DateTime.TryParse(worksheet1.Cells[row, 38].Value?.ToString(), out DateTime parsedDate) ? parsedDate : (DateTime?)null;
                                    existingRecord.BatchDeadline = DateTime.TryParse(worksheet1.Cells[row, 38].Value?.ToString(), out DateTime parsedDate1)
                                        ? parsedDate1.AddDays(60)
                                        : (DateTime?)null;
                                    existingRecord.Status = "True";


                                    if (!(existingRecord.OS_Billing == null || existingRecord.OS_Billing == "-" || existingRecord.OS_Billing.Trim() == "0") && !(existingRecord.ExpectedRenewalFee == null || existingRecord.ExpectedRenewalFee == "-" || existingRecord.ExpectedRenewalFee.Trim() == "0"))
                                    {
                                        existingRecord.Segments = "Bounced Cheque and Renewal";
                                    }
                                    else if (!(existingRecord.OS_Billing == null || existingRecord.OS_Billing == "-" || existingRecord.OS_Billing.Trim() == "0"))
                                    {
                                        existingRecord.Segments = "Bounced Cheque";
                                    }
                                    else if (!(existingRecord.ExpectedRenewalFee == null || existingRecord.ExpectedRenewalFee == "-" || existingRecord.ExpectedRenewalFee.Trim() == "0"))
                                    {
                                        existingRecord.Segments = "Renewal";
                                    }
                                    else
                                    {
                                        existingRecord.Segments = " ";
                                    }

                                    dbSheet1.SaveChanges();

                                    int recordId = existingRecord.Id; // Get the auto-generated Id after saving

                                    InsertMobileNo(dbSheet1, recordId, existingRecord.Mobile1);
                                    InsertMobileNo(dbSheet1, recordId, existingRecord.Mobile2);
                                    InsertMobileNo(dbSheet1, recordId, existingRecord.Mobile3);
                                    InsertMobileNo(dbSheet1, recordId, existingRecord.Mobile4);

                                    InsertEmailId(dbSheet1, recordId, existingRecord.Email_1);
                                    InsertEmailId(dbSheet1, recordId, existingRecord.Email_2);
                                    InsertEmailId(dbSheet1, recordId, existingRecord.Email_3);
                                }
                            }
                            else
                            {
                                var caseEntity1 = new RecordData
                                {
                                    AccountNo = worksheet1.Cells[row, 1].Value?.ToString(),
                                    CustomerName = worksheet1.Cells[row, 2].Value?.ToString(),
                                    Contact_Person = worksheet1.Cells[row, 3].Value?.ToString(),
                                    Nationality = worksheet1.Cells[row, 4].Value?.ToString(),
                                    Mobile1 = worksheet1.Cells[row, 5].Value?.ToString(),
                                    Mobile2 = worksheet1.Cells[row, 6].Value?.ToString(),
                                    Mobile3 = worksheet1.Cells[row, 7].Value?.ToString(),
                                    Mobile4 = worksheet1.Cells[row, 8].Value?.ToString(),
                                    Email_1 = worksheet1.Cells[row, 9].Value?.ToString(),
                                    Email_2 = worksheet1.Cells[row, 10].Value?.ToString(),
                                    Email_3 = worksheet1.Cells[row, 11].Value?.ToString(),
                                    TenacyFacilityType = worksheet1.Cells[row, 12].Value?.ToString(),
                                    License_expiry = worksheet1.Cells[row, 13].Value?.ToString(),
                                    ExpectedRenewalFee = worksheet1.Cells[row, 14].Value?.ToString(),
                                    SRNumber = worksheet1.Cells[row, 15].Value?.ToString(),
                                    DeRegFee = worksheet1.Cells[row, 16].Value?.ToString(),
                                    BCheque = worksheet1.Cells[row, 17].Value?.ToString(),
                                    BCheque_P = worksheet1.Cells[row, 18].Value?.ToString(),
                                    IPTelephone_Billing = worksheet1.Cells[row, 19].Value?.ToString(),
                                    Utility_Billing = worksheet1.Cells[row, 20].Value?.ToString(),
                                    Others = worksheet1.Cells[row, 21].Value?.ToString(),
                                    OS_Billing = worksheet1.Cells[row, 22].Value?.ToString(),
                                    CloseAccount = worksheet1.Cells[row, 23].Value?.ToString(),
                                    DormantAccount = worksheet1.Cells[row, 24].Value?.ToString(),
                                    InsufficientFunds = worksheet1.Cells[row, 25].Value?.ToString(),
                                    OtherReason = worksheet1.Cells[row, 26].Value?.ToString(),
                                    SignatureIrregular = worksheet1.Cells[row, 27].Value?.ToString(),
                                    TechnicalReason = worksheet1.Cells[row, 28].Value?.ToString(),
                                    BOthers = worksheet1.Cells[row, 29].Value?.ToString(),
                                    EmployeeVisaQuota = worksheet1.Cells[row, 30].Value?.ToString(),
                                    EmployeeVisaUtilized = worksheet1.Cells[row, 31].Value?.ToString(),
                                    ProjectBundleName = worksheet1.Cells[row, 32].Value?.ToString(),
                                    LicenseType = worksheet1.Cells[row, 33].Value?.ToString(),
                                    FacilityType = worksheet1.Cells[row, 34].Value?.ToString(),
                                    NoYears = worksheet1.Cells[row, 35].Value?.ToString(),
                                    DerbyBatch = worksheet1.Cells[row, 36].Value?.ToString(),
                                    Agent = worksheet1.Cells[row, 37].Value?.ToString(),
                                    Status = "True",
                                    BatchDate = DateTime.TryParse(worksheet1.Cells[row, 38].Value?.ToString(), out DateTime parsedDate) ? parsedDate : (DateTime?)null,
                                    BatchDeadline = DateTime.TryParse(worksheet1.Cells[row, 38].Value?.ToString(), out DateTime parsedDate1)
                                        ? parsedDate1.AddDays(60)
                                        : (DateTime?)null
                            };

                                if (!(caseEntity1.OS_Billing == null || caseEntity1.OS_Billing == "-" || caseEntity1.OS_Billing.Trim() == "0") && !(caseEntity1.ExpectedRenewalFee == null || caseEntity1.ExpectedRenewalFee == "-" || caseEntity1.ExpectedRenewalFee.Trim() == "0"))
                                {
                                    caseEntity1.Segments = "Bounced Cheque and Renewal";
                                }
                                else if (!(caseEntity1.OS_Billing == null || caseEntity1.OS_Billing == "-" || caseEntity1.OS_Billing.Trim() == "0"))
                                {
                                    caseEntity1.Segments = "Bounced Cheque";
                                }
                                else if (!(caseEntity1.ExpectedRenewalFee == null || caseEntity1.ExpectedRenewalFee == "-" || caseEntity1.ExpectedRenewalFee.Trim() == "0"))
                                {
                                    caseEntity1.Segments = "Renewal";
                                }
                                else
                                {
                                    caseEntity1.Segments = " ";
                                }

                                dbSheet1.RecordDatas.Add(caseEntity1);
                                dbSheet1.SaveChanges();

                                int recordId = caseEntity1.Id; // Get the auto-generated Id after saving

                                InsertMobileNo(dbSheet1, recordId, caseEntity1.Mobile1);
                                InsertMobileNo(dbSheet1, recordId, caseEntity1.Mobile2);
                                InsertMobileNo(dbSheet1, recordId, caseEntity1.Mobile3);
                                InsertMobileNo(dbSheet1, recordId, caseEntity1.Mobile4);

                                InsertEmailId(dbSheet1, recordId, caseEntity1.Email_1);
                                InsertEmailId(dbSheet1, recordId, caseEntity1.Email_2);
                                InsertEmailId(dbSheet1, recordId, caseEntity1.Email_3);
                            }
                            // Insert mobile numbers into MobileNos table



                        }

                       


                        // Process the second sheet
                        ExcelWorksheet worksheet2 = package.Workbook.Worksheets["BC Details"];
                        int rowCount2 = worksheet2.Dimension.Rows;


                        for (int row = 2; row <= rowCount2; row++) // Start from row 2 to skip headers
                        {
                            var accountNo = worksheet2.Cells[row, 1].Value?.ToString();

                            // Find all records with the same accountNo and remove them
                            var existingRecords = dbSheet2.BouncedRecords.Where(rd => rd.AccountNo == accountNo).ToList();
                            foreach (var record in existingRecords)
                            {
                                dbSheet2.BouncedRecords.Remove(record);
                            }

                            // Create and add the new record
                            var newRecord = new BouncedRecord
                            {
                                AccountNo = worksheet2.Cells[row, 1].Value?.ToString(),
                                ChequeNumber = worksheet2.Cells[row, 2].Value?.ToString(),
                                DateBounced = (DateTime?)worksheet2.Cells[row, 3].Value,
                                TotalAmount = worksheet2.Cells[row, 4].Value?.ToString(),
                                ReasonCode = worksheet2.Cells[row, 5].Value?.ToString(),
                                Text = worksheet2.Cells[row, 6].Value?.ToString(),
                            };

                            dbSheet2.BouncedRecords.Add(newRecord);
                        }



                        //for (int row = 2; row <= rowCount2; row++) // Start from row 2 to skip headers
                        //{

                        //    var caseEntity2 = new BouncedRecord // Assuming you have a different model for the second sheet
                        //    {
                        //        AccountNo = worksheet2.Cells[row, 1].Value?.ToString(),
                        //        ChequeNumber = worksheet2.Cells[row, 2].Value?.ToString(),
                        //        DateBounced = (DateTime?)worksheet2.Cells[row, 3].Value,
                        //        TotalAmount = worksheet2.Cells[row, 4].Value?.ToString(),
                        //        ReasonCode = worksheet2.Cells[row, 5].Value?.ToString(),
                        //        Text = worksheet2.Cells[row, 6].Value?.ToString(),
                        //        //ChequeDate = (DateTime)worksheet2.Cells[row, 6].Value,

                        //    };

                        //    dbSheet2.BouncedRecords.Add(caseEntity2);
                        //}



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

   

        //void InsertMobileNo(CPVDBEntities context, int recordId, string mobileNo)
        //{
        //    if (!string.IsNullOrEmpty(mobileNo))
        //    {
        //        var mobileRecord = new MobileNo
        //        {
        //            RecordId = recordId,
        //            Numbers = mobileNo
        //        };

        //        context.Set<MobileNo>().Add(mobileRecord);
        //        context.SaveChanges();
        //    }
        //}

        void InsertMobileNo(CPVDBEntities context, int recordId, string mobileNo)
        {
            if (!string.IsNullOrEmpty(mobileNo))
            {
                var existingMobile = context.MobileNos.FirstOrDefault(m => m.RecordId == recordId && m.Numbers == mobileNo);
                if (existingMobile != null)
                {
                    existingMobile.Numbers = mobileNo; // Update the number if it exists
                }
                else
                {
                    var mobileRecord = new MobileNo
                    {
                        RecordId = recordId,
                        Numbers = mobileNo
                    };

                    context.Set<MobileNo>().Add(mobileRecord); // Add a new number if it does not exist
                }
                context.SaveChanges();
            }
        }


        //void InsertEmailId(CPVDBEntities context, int recordId, string emailId)
        //{
        //    if (!string.IsNullOrEmpty(emailId))
        //    {
        //        var emailRecord = new EmailId
        //        {
        //            RecordId = recordId,
        //            Emails = emailId
        //        };

        //        context.Set<EmailId>().Add(emailRecord);
        //        context.SaveChanges();
        //    }
        //}

        void InsertEmailId(CPVDBEntities context, int recordId, string emailId)
        {
            if (!string.IsNullOrEmpty(emailId))
            {
                var existingEmail = context.EmailIds.SingleOrDefault(e => e.RecordId == recordId && e.Emails == emailId);
                if (existingEmail != null)
                {
                    existingEmail.Emails = emailId; // Update the email if it exists
                }
                else
                {
                    var emailRecord = new EmailId
                    {
                        RecordId = recordId,
                        Emails = emailId
                    };

                    context.Set<EmailId>().Add(emailRecord); // Add a new email if it does not exist
                }
                context.SaveChanges();
            }
        }


        public static List<string> GetAllUsers()
        {
            using (CPVDBEntities db = new CPVDBEntities())
            {
                List<string> users = db.AspNetUsers
                 .Where(w => w.UserRole == "Agent")
                 .Select(s => s.UserName)
                 .ToList();

                return users;
            }
        }


        [Authorize]
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

        public ActionResult AllocationReport()
        {
            CPVDBEntities db = new CPVDBEntities();
            RecordData recordData = new RecordData();

            var AgentNames = db.AspNetUsers
                         .Where(r => r.UserRole == "Agent")
                         .Select(r => r.UserName)
                         .Distinct()
                         .ToList();
            ViewBag.AgentNames = AgentNames;

            List<string> allbatches = db.RecordDatas
                  .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                  .Select(r => r.DerbyBatch)
                  .Distinct()
                  .OrderBy(DerbyBatch => DerbyBatch)
                  .ToList();

            ViewBag.allbatches = allbatches;

            return View(recordData);
        }

        public ActionResult AllocationReportDownload(string Users, string CallType, string Disposition, string SubDisposition, string DerbyBatch, DateTime? CallbackTime = null, DateTime? specificDate = null, DateTime? endDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = db.EventTables.AsQueryable();

            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            List<EventTable> eventTables = new List<EventTable>();
            List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            bouncedRecords = db.BouncedRecords.ToList();



            var key = "";

            if (DerbyBatch != "")
            {
                var accountNumbers = db.RecordDatas.Where(rd => rd.DerbyBatch == DerbyBatch).Select(rd => rd.AccountNo).ToList();
                query = query.Where(et => accountNumbers.Contains(et.AccountNo));

            }

            if (Users != "")
            {
                query = query.Where(et => et.Agent == Users);
                var hasRows = query.Any();

            }

            if (CallType != "-")
            {
                query = query.Where(et => et.CallType == CallType);
            }

            if (Disposition != "-")
            {
                query = query.Where(et => et.Dispo == Disposition);
            }

            if (SubDisposition != "-")
            {
                query = query.Where(et => et.SubDispo == SubDisposition);
            }

            if (CallbackTime != null)
            {
                query = query.Where(et => et.CallbackTime == CallbackTime);
            }

            if (specificDate != null)
            {
                query = query.Where(et => et.Datetime >= specificDate);
            }

            if (endDate != null)
            {
                query = query.Where(et => et.Datetime <= endDate);
            }

            var caseTables = query;

            var accountNos = caseTables.Select(ct => ct.AccountNo).ToList();

            //recordDatas = accountNos
            //   .SelectMany(accountNo => db.RecordDatas
            //       .Where(rd => rd.AccountNo == accountNo))
            //   .ToList();

            //recordDatas = db.RecordDatas.Where(rd => rd.AccountNo == key).ToList();


            var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing",
                "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
                "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee",
                 "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
                "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            var header1 = new List<string>()
            {
                "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments", "Segments","Datetime"
            };

            var header2 = new List<string>()
            {
                "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments", "Segments","DateTime"
            };

            var bouncedDetailHeaders = new List<string>()
            {
                 "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", " ChequeDate","TotalAmount"
            };


            using (var package = new ExcelPackage())
            {
                int col = 1;
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var combinedList = header.Concat(header2).ToList();
                List<string> accountNoList = new List<string>();


                foreach (var headerName in combinedList)
                {
                    worksheet.Cells[1, col++].Value = headerName.ToString();
                }



                int row = 2;

                foreach (var items in caseTables)
                {
                    col = 31;
                    foreach (var headName in header1)
                    {
                        var property = typeof(EventTable).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

                        if (property != null && items != null)
                        {

                            object value = property.GetValue(items);

                            if (property != null && property.Name == "AccountNo")
                            {
                                key = value.ToString();
                                accountNoList.Add(value?.ToString());
                            }

                            if (value != null && value.ToString() == "1/1/2000 12:00:00 AM")
                            {

                                worksheet.Cells[row, col++].Value = "-";
                            }
                            else
                            {
                                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                            }
                        }
                    }

                    recordDatas = db.RecordDatas.Where(rd => rd.AccountNo == key).ToList();
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
                                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                            }
                        }
                    }

                    row++;
                }


                var bouncedWorksheet = package.Workbook.Worksheets.Add("BouncedDetails");
                col = 1;
                foreach (var headerName in bouncedDetailHeaders)
                {
                    bouncedWorksheet.Cells[1, col++].Value = headerName.ToString();
                }


                row = 2;

                foreach (var bouncedDetail in bouncedRecords)
                {
                    col = 1;
                    foreach (var headerName in bouncedDetailHeaders)
                    {


                        var property = typeof(BouncedRecord).GetProperty(headerName, BindingFlags.Public | BindingFlags.Instance);

                        if (property != null && bouncedDetail != null)
                        {
                            object value = property.GetValue(bouncedDetail);

                            if (property.PropertyType == typeof(DateTime))
                            {
                                bouncedWorksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            else
                            {
                                bouncedWorksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
                            }
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


        public ActionResult DownloadCases(string Users, string CallType,string Disposition,string SubDisposition, string DerbyBatch, DateTime? CallbackTime = null, DateTime? specificDate = null, DateTime? endDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = db.EventTables.AsQueryable();
           
            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            List<EventTable> eventTables = new List<EventTable>();
            List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            bouncedRecords = db.BouncedRecords.ToList();  
            


            var key = "";
                
            if(DerbyBatch != "")
            {
                var accountNumbers = db.RecordDatas.Where(rd => rd.DerbyBatch == DerbyBatch).Select(rd => rd.AccountNo).ToList();
                query = query.Where(et => accountNumbers.Contains(et.AccountNo));
               
            }

            if (Users != "")
            {
                query = query.Where(et => et.Agent == Users);
                var hasRows = query.Any();

            }

            if (CallType != "-")
            {
                query = query.Where(et => et.CallType == CallType);
            }

            if (Disposition != "-")
            {
                query = query.Where(et => et.Dispo == Disposition);
            }

            if (SubDisposition != "-")
            {
                query = query.Where(et => et.SubDispo == SubDisposition);
            }

            if (CallbackTime != null)
            {
                query = query.Where(et => et.CallbackTime == CallbackTime);
            } 

            if (specificDate != null)
            {
                query = query.Where(et => et.Datetime >= specificDate);
            }

            if (endDate != null)
            {
                query = query.Where(et => et.Datetime <= endDate);
            }

            var caseTables = query;

            var accountNos = caseTables.Select(ct => ct.AccountNo).ToList();

            //recordDatas = accountNos
            //   .SelectMany(accountNo => db.RecordDatas
            //       .Where(rd => rd.AccountNo == accountNo))
            //   .ToList();

            //recordDatas = db.RecordDatas.Where(rd => rd.AccountNo == key).ToList();


            var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing",
                "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
                "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee",
                 "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
                "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            var header1 = new List<string>()
            {
                "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments", "Segments","Datetime"
            };

            var header2 = new List<string>()
            {
                "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments", "Segments","DateTime"
            };

            var bouncedDetailHeaders = new List<string>()
            {
                 "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", " ChequeDate","TotalAmount"
            };


            using (var package = new ExcelPackage())
            {
                int col = 1;
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var combinedList = header.Concat(header2).ToList();
                List<string> accountNoList = new List<string>();


                foreach (var headerName in combinedList)
                {
                    worksheet.Cells[1, col++].Value = headerName.ToString();
                }



                int row = 2;

                foreach (var items in caseTables)
                {
                    col = 31;
                    foreach (var headName in header1)
                    {
                        var property = typeof(EventTable).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

                        if (property != null && items != null)
                        {

                            object value = property.GetValue(items);

                            if (property != null && property.Name == "AccountNo")
                            {
                                key = value.ToString();
                                accountNoList.Add(value?.ToString());
                            }

                            if (value != null && value.ToString() == "1/1/2000 12:00:00 AM")
                            {

                                worksheet.Cells[row, col++].Value = "-";
                            }
                            else
                            {
                                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                            }
                        }
                    }

                    recordDatas = db.RecordDatas.Where(rd => rd.AccountNo == key).ToList();
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
                                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                            }
                        }
                    }

                    row++;
                }

               
                var bouncedWorksheet = package.Workbook.Worksheets.Add("BouncedDetails");
                col = 1;
                foreach (var headerName in bouncedDetailHeaders)
                {
                    bouncedWorksheet.Cells[1, col++].Value = headerName.ToString();
                }


                row = 2;

                foreach (var bouncedDetail in bouncedRecords)
                {
                    col = 1;
                    foreach (var headerName in bouncedDetailHeaders)
                    {
                        
                        
                        var property = typeof(BouncedRecord).GetProperty(headerName, BindingFlags.Public | BindingFlags.Instance);

                        if (property != null && bouncedDetail != null)
                        {
                            object value = property.GetValue(bouncedDetail);

                            if (property.PropertyType == typeof(DateTime))
                            {
                                bouncedWorksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            else
                            {
                                bouncedWorksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
                            }
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

        public ActionResult DownloadRecordCases(string Users, string CallType, string Disposition, string SubDisposition, string DerbyBatch, string Status, DateTime? CallbackTime = null, DateTime? specificDate = null, DateTime? endDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = db.RecordDatas.AsQueryable();

            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            List<EventTable> eventTables = new List<EventTable>();
            List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            bouncedRecords = db.BouncedRecords.ToList();

            var key = "";

            if(DerbyBatch != "")
            {
                query = query.Where(et => et.DerbyBatch == DerbyBatch);
               
            }

            if (Users != "")
            {
                query = query.Where(et => et.Agent == Users);
                var hasRows = query.Any();

            }

            if (CallType != "-")
            {
                query = query.Where(et => et.CallType == CallType);
            }

            if (Disposition != "-")
            {
                query = query.Where(et => et.Disposition == Disposition);
            }

            if (SubDisposition != "-")
            {
                query = query.Where(et => et.SubDisposition == SubDisposition);
            }

            if (CallbackTime != null)
            {
                query = query.Where(et => et.CallbackTime == CallbackTime);
            }

            if (specificDate != null)
            {
                query = query.Where(et => et.ModifiedDate >= specificDate);
            }

            if (endDate != null)
            {
                query = query.Where(et => et.ModifiedDate <= endDate.Value.Date);
            }

            if(Status == "Active")
            {
                query = query.Where(et => et.Status == "True");
            }
            else if(Status == "Inactive")
            {
                query = query.Where(et => et.Status == "False");
            }

            var caseTables = query;

            var accountNos = caseTables.Select(ct => ct.AccountNo).ToList();

            //var matchingEvents = db.EventTables
            //.Where(et => accountNos.Contains(et.AccountNo))
            //.ToList();

            //var latestEvents = matchingEvents
            //.GroupBy(et => et.AccountNo)
            //.Select(g => g.OrderByDescending(et => et.Datetime).FirstOrDefault())
            //.ToList();

            //var latestEvents = db.EventTables
            //        .Where(et => accountNos.Contains(et.AccountNo))
            //       .GroupBy(et => et.AccountNo)
            //       .Select(g => g.OrderByDescending(et => et.Datetime).FirstOrDefault())
            //       .ToList();

            //eventTables = db.EventTables
            //                 //.Where(rd => rd.AccountNo == key)
            //                .Where(rd => accountNos.Contains(rd.AccountNo))
            //                 .GroupBy(rd => rd.AccountNo)
            //                 .Select(g => g.OrderByDescending(rd => rd.Datetime).FirstOrDefault())
            //                 .ToList();

            //eventTables = db.EventTables
            //       .Where(rd => accountNos.Contains(rd.AccountNo))
            //       .OrderByDescending(rd => rd.Datetime) // Optional: if you want to keep the events ordered by Datetime
            //       .ToList();

             //eventTables = db.EventTables
             //       .Where(rd => accountNos.Contains(rd.AccountNo))
             //       .GroupBy(rd => rd.AccountNo)
             //       .Select(g => g.OrderByDescending(rd => rd.Datetime).FirstOrDefault())
             //       .ToList();

             eventTables = accountNos
                .Select(accountNo => db.EventTables
                    .Where(rd => rd.AccountNo == accountNo)
                    .OrderByDescending(rd => rd.Datetime)
                    .FirstOrDefault())
                .ToList();


            var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing",
                "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
                "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee",
                 "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
                "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            var header1 = new List<string>()
            {
                "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments"
            };

            var header2 = new List<string>()
            {
                "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments","Segments"
            };

            var bouncedDetailHeaders = new List<string>()
            {
                 "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", "ChequeDate","TotalAmount"
            };

            var ReqAsked = new List<string>()
            {
                "Segments"
            };

            using (var package = new ExcelPackage())
            {
                int col = 1;
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var combinedList = header.Concat(header2).ToList();
                List<string> accountNoList = new List<string>();


                foreach (var headerName in combinedList)
                {
                    worksheet.Cells[1, col++].Value = headerName.ToString();
                }

               
                int row = 2;

                foreach (var items in caseTables)
                {
                    col = 1;
                    foreach (var headName in header)
                    {
                       var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);
                        
                        object value = property.GetValue(items);

                        if (property != null && property.Name == "AccountNo")
                        {
                            key = value.ToString();
                            accountNoList.Add(value?.ToString());
                        }

                        if (value != null && value.ToString() == "1/1/2000 12:00:00 AM")
                        {

                            worksheet.Cells[row, col++].Value = "-";
                        }
                        else
                        {
                            worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                        }
                    }
                    //eventTables = db.EventTables
                    //                 .Where(rd => rd.AccountNo == key)
                    //                 .GroupBy(rd => rd.AccountNo)
                    //                 .Select(g => g.OrderByDescending(rd => rd.Datetime).FirstOrDefault())
                    //                 .ToList();




                    col = 39;
                    foreach (var headName in ReqAsked)
                    {
                        var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);
                        object value = property.GetValue(items);


                        worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";

                    }

                    row++;
                }

                row = 2;
                foreach (var itemcase1 in eventTables)
                {
                    col = 31;   
                    foreach (var headName1 in header1)
                    {
                        var property = typeof(EventTable).GetProperty(headName1, BindingFlags.Public | BindingFlags.Instance);

                        if (property != null && itemcase1 != null)
                        {
                            var value1 = property.GetValue(itemcase1);

                            if (property.PropertyType == typeof(DateTime))
                            {
                                worksheet.Cells[row, col++].Value = ((DateTime)value1).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            else if (value1 != null && value1.ToString() == "1/1/2000 12:00:00 AM")
                            {

                                worksheet.Cells[row, col++].Value = "-";
                            }
                            else
                            {
                                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value1?.ToString()) ? value1.ToString() : "-";
                            }
                        }

                    }
                    row++;

                }



                var bouncedWorksheet = package.Workbook.Worksheets.Add("BouncedDetails");
                col = 1;
                foreach (var headerName in bouncedDetailHeaders)
                {
                    bouncedWorksheet.Cells[1, col++].Value = headerName.ToString();
                }
                row = 2;
                foreach (var bouncedDetail in bouncedRecords)
                {
                    col = 1;
                    foreach (var headerName in bouncedDetailHeaders)
                    {
                        var property = typeof(BouncedRecord).GetProperty(headerName, BindingFlags.Public | BindingFlags.Instance);
                        object value = property.GetValue(bouncedDetail);

                        if (property.PropertyType == typeof(DateTime))
                        {
                            bouncedWorksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            bouncedWorksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
                        }
                    }
                    row++;
                }

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                string batchName = string.IsNullOrEmpty(DerbyBatch) ? "_AllBatch_" : "_"+DerbyBatch+"_";
                string fileName = "Caselist" + batchName + DateTime.Today.ToString("yyyy-MM-dd") + ".xlsx";
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName);



                //Response.AddHeader("content-disposition", "attachment;filename=Caselist" + DerbyBatch + (DateTime.Today.ToString("yyyy-MM-dd")) + ".xlsx");

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