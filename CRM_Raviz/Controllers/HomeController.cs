﻿using CRM_Raviz.Models;
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
            RecordData recordData = new RecordData();

            var AgentNames = db.AspNetUsers
                         .Where(r => r.UserRole == "Agent")
                         .Select(r => r.UserName)
                         .Distinct()
                         .ToList();

            // Pass the dispositionValues to the view
            ViewBag.AgentNames = AgentNames;

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

            // Pass the dispositionValues to the view
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

            // Return any response if needed
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





        //public ActionResult Index()
        //{
        //    DateTime today = DateTime.Today;

        //    CPVDBEntities db = new CPVDBEntities();

        //    // Query to get the list of all agents
        //    var allAgents = db.RecordDatas.Select(a => a.Agent).Distinct().ToList();


        //    var startOfDay = today.Date;
        //    var endOfDay = today.Date.AddDays(1).AddTicks(-1); // Set end of day to 23:59:59.999

        //    var agentCasesCount = db.RecordDatas
        //                            .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfDay && r.ModifiedDate <= endOfDay)
        //                            .GroupBy(r => r.Agent)
        //                            .Select(g => new { Agent = g.Key, CasesCount = g.Count() })
        //                            .ToList();



        //    var agentsWithCasesCount = allAgents
        //                                .Select(agent => new
        //                                {
        //                                    Agent = agent,
        //                                    CasesCount = agentCasesCount.FirstOrDefault(ac => ac.Agent == agent)?.CasesCount ?? 0
        //                                })
        //                                .ToList();

        //    ViewBag.AgentsWithCasesCount = agentsWithCasesCount; // Pass the list of agents with counts to the view using ViewBag

        //    return View();
        //}



        //public ActionResult Index()
        //{
        //    DateTime today = DateTime.Today;

        //    using (CPVDBEntities db = new CPVDBEntities())
        //    {
        //        // Query to get the list of all agents
        //        var allAgents = db.RecordDatas.Select(a => a.Agent).Distinct().ToList();

        //        var startOfDay = today.Date;
        //        var endOfDay = today.Date.AddDays(1).AddTicks(-1); // Set end of day to 23:59:59.999

        //        // Query to get the count of cases for each agent for today's date
        //        var agentCasesCount = db.RecordDatas
        //                                .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfDay && r.ModifiedDate <= endOfDay)
        //                                .GroupBy(r => r.Agent)
        //                                .Select(g => new { Agent = g.Key, CasesCount = g.Count() })
        //                                .ToList();

        //        // Combine all agents with today's cases count

        //        var agentsWithCasesCount = allAgents
        //                                    .Select(agent => new
        //                                    {
        //                                        Agent = agent,
        //                                        CasesCount = agentCasesCount.FirstOrDefault(ac => ac.Agent == agent)?.CasesCount ?? 0
        //                                    })
        //                                    .ToList();

        //        ViewBag.AgentsWithCasesCount = agentsWithCasesCount; // Pass the list of agents with counts to the view using ViewBag
        //    }

        //    return View();
        //}

        //public ActionResult Index()
        //{
        //    DateTime today = DateTime.Today;
        //    using (CPVDBEntities db = new CPVDBEntities())
        //    {
        //        var allAgents = db.RecordDatas.Select(a => a.Agent).Distinct().ToList();

        //        var startOfDay = today.Date;
        //        var endOfDay = today.Date.AddDays(1).AddTicks(-1); 


        //        var agentCasesCount = db.RecordDatas
        //            .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfDay && r.ModifiedDate <= endOfDay)
        //            .GroupBy(r => r.Agent)
        //            .Select(g => new { Agent = g.Key, CasesCount = g.Count() })
        //            .ToList();

        //        // Combine all agents with today's cases count
        //        var agentsWithCasesCount = allAgents
        //            .Select(agent => new AgentCasesViewModel
        //            {
        //                Agent = agent,
        //                CasesCount = agentCasesCount.FirstOrDefault(ac => ac.Agent == agent)?.CasesCount ?? 0
        //            })
        //            .ToList();

        //        return View(agentsWithCasesCount);
        //    }
        //}

        //public ActionResult Index()                                                                 
        //{
        //    DateTime today = DateTime.Today;
        //    using (CPVDBEntities db = new CPVDBEntities())
        //    {
        //        var allAgents = db.RecordDatas
        //            .Where(a => a.Agent != null)
        //            .Select(a => a.Agent)
        //            .Distinct()
        //            .ToList();
                        
        //        var startOfDay = today.Date;
        //        var endOfDay = today.Date.AddDays(1).AddTicks(-1);

        //        var startOfMonth = new DateTime(today.Year, today.Month, 1);
        //        var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);


        //        int casesCountToday = db.RecordDatas
        //        .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfDay && r.ModifiedDate <= endOfDay)
        //        .Count();

        //        int casesCountThisMonth = db.RecordDatas
        //        .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfMonth && r.ModifiedDate <= endOfMonth)
        //        .Count();

        //        int totalCases = db.RecordDatas
        //        .Count();

        //         int totalCasesnotworked = db.RecordDatas
        //         .Where(r => r.ModifiedDate == null )
        //         .Count();

                

        //         int callbackCountThisMonth = db.RecordDatas
        //        .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
        //        .Count();

        //        var segmentCounts = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments))
        //        .GroupBy(r => r.Segments)
        //        .OrderBy(g => g.Key) // Order groups by segment name
        //        .Select(g => g.Count())
        //        .ToList();

        //        //var segmentNotWorked = db.RecordDatas
        //        //.Where(r => !string.IsNullOrEmpty(r.Segments) && r.ModifiedDate == null)
        //        //.GroupBy(r => r.Segments)
        //        //.OrderBy(g => g.Key) 
        //        //.Select(g => g.Count())
        //        //.ToList();


        //        // Define your list of segments to check
        //        //List<string> allSegmentsToCheck = new List<string> { "Segment1", "Segment2", /* add more segments as needed */ };


        //        List<string> allSegments = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments))
        //        .Select(r => r.Segments)
        //        .Distinct()
        //        .OrderBy(segment => segment)
        //        .ToList();


        //        List<int> segmentNotWorked = allSegments
        //        .Select(segment => db.RecordDatas
        //            .Count(r => r.Segments == segment && r.ModifiedDate == null))
        //        .ToList();

        //        // Query the database

        //        //var segmentNotWorked = allSegments
        //        //    .Select(segment => new
        //        //    {
        //        //        Segment = segment,
        //        //        Count = db.RecordDatas
        //        //            .Where(r => r.Segments == segment && r.ModifiedDate == null)
        //        //            .Count()
        //        //    })
        //        //    .OrderBy(result => result.Segment)
        //        //    .ToList();



        //        var distinctSegments = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments))
        //        .Select(r => r.Segments)
        //        .OrderBy(segment => segment) // Order by segment name
        //        .Distinct()
        //        .ToList();


        //        var segment2Counts = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments))
        //        .GroupBy(r => r.DerbyBatch)
        //        .OrderBy(g => g.Key) // Order groups by segment name
        //        .Select(g => g.Count())
        //        .ToList();

        //        var BatchNotWorked = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments) && r.ModifiedDate == null)
        //        .GroupBy(r => r.DerbyBatch)
        //        .OrderBy(g => g.Key) // Order groups by segment name
        //        .Select(g => g.Count())
        //        .ToList();

        //        var BatchWorked = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments) && r.ModifiedDate != null)
        //        .GroupBy(r => r.DerbyBatch)
        //        .OrderBy(g => g.Key) // Order groups by segment name
        //        .Select(g => g.Count())
        //        .ToList();


        //        var distinctSegments2 = db.RecordDatas
        //        .Where(r => !string.IsNullOrEmpty(r.Segments))
        //        .Select(r => r.DerbyBatch)
        //        .OrderBy(DerbyBatch => DerbyBatch) // Order by segment name
        //        .Distinct()
        //        .ToList();





        //        ViewBag.totalCasesnotworked = totalCasesnotworked;
        //        ViewBag.totalCases = totalCases;

        //        ViewBag.segmentNotWorked = segmentNotWorked;
        //        ViewBag.BatchNotWorked = BatchNotWorked;
        //        ViewBag.BatchWorked = BatchWorked;
        //        ViewBag.segment2Counts = segment2Counts;
        //        ViewBag.distinctSegments2 = distinctSegments2;
        //        ViewBag.distinctSegments = distinctSegments;
        //        ViewBag.segmentCounts = segmentCounts;
        //        ViewBag.recentEvents = casesCountToday;
        //        ViewBag.callbackCountThisMonth = callbackCountThisMonth;
        //        ViewBag.CasesCountThisMonth = casesCountThisMonth;

        //        var agentCasesCountToday = db.RecordDatas
        //            .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfDay && r.ModifiedDate <= endOfDay)
        //            .GroupBy(r => r.Agent)
        //            .Select(g => new { Agent = g.Key, CasesCountToday = g.Count() })
        //            .ToList();



        //        var agentCasesCountThisMonth = db.RecordDatas
        //            .Where(r => r.ModifiedDate.HasValue && r.ModifiedDate >= startOfMonth && r.ModifiedDate <= endOfMonth)
        //            .GroupBy(r => r.Agent)
        //            .Select(g => new { Agent = g.Key, CasesCountThisMonth = g.Count() })
        //            .ToList();

        //        var agentsWithCasesCount = allAgents
        //            .Select(agent => new AgentCasesViewModel
        //            {
        //                Agent = agent,
        //                CasesCountToday = agentCasesCountToday.FirstOrDefault(ac => ac.Agent == agent)?.CasesCountToday ?? 0,
        //                CasesCountThisMonth = agentCasesCountThisMonth.FirstOrDefault(ac => ac.Agent == agent)?.CasesCountThisMonth ?? 0
        //            })
        //            .ToList();

        //        var agentCallBackCountToday = db.RecordDatas
        //            .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
        //            .GroupBy(r => r.Agent)
        //            .Select(g => new { Agent = g.Key, CallBackCountToday = g.Count() })
        //            .ToList();



        //        var agentCallBackCountPrev = db.RecordDatas
        //            .Where(r => r.CallbackTime.HasValue && r.CallbackTime <= startOfDay && r.CallbackTime != new DateTime(2000, 1, 1))
        //            .GroupBy(r => r.Agent)
        //            .Select(g => new { Agent = g.Key, CallBackCountPrev = g.Count() })
        //            .ToList();

        //        var agentCallBackCount = allAgents
        //            .Select(agent => new AgentCasesViewModel
        //            {
        //                Agent = agent,
        //                CallBackCountToday = agentCallBackCountToday.FirstOrDefault(ac => ac.Agent == agent)?.CallBackCountToday ?? 0,
        //                CallBackCountPrev = agentCallBackCountPrev.FirstOrDefault(ac => ac.Agent == agent)?.CallBackCountPrev ?? 0
        //            })
        //            .ToList();


        //        ViewBag.AgentCallBackCount = agentCallBackCount;

        //        return View(agentsWithCasesCount);
        //    }
        //}

        public ActionResult Index(string filter)
        {
            DateTime today = DateTime.Today;

            using (CPVDBEntities db = new CPVDBEntities())
            {
                var records = db.RecordDatas.ToList();
                string currentUser = User.Identity.GetUserName();

                IQueryable<RecordData> recordDatasQuery = db.RecordDatas;
                IQueryable<EventTable> eventTableQuery = db.EventTables;

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

                    recordDatasQuery = db.RecordDatas.Where(r => r.ModifiedAgent == filter);
                    eventTableQuery = db.EventTables.Where(r => r.Agent == filter);

                    allAgents = db.AspNetUsers
                      .Where(r => r.UserRole == "Agent" && r.UserName == filter)
                      .Select(r => r.UserName)
                      .Distinct()
                      .ToList();

                    ViewBag.Filter = filter;
                }
               
               
                if (User.IsInRole("Agent"))
                {
                    recordDatasQuery = recordDatasQuery.Where(r => r.ModifiedAgent == currentUser);
                    eventTableQuery = eventTableQuery.Where(r => r.Agent == currentUser);

                     allAgents = db.AspNetUsers
                       .Where(r => r.UserRole == "Agent" && r.UserName == currentUser)
                       .Select(r => r.UserName)
                       .Distinct()
                       .ToList();

                    ViewBag.Filter = currentUser;
                }
              
              
                    
            


                var startOfDay = today.Date;
                var endOfDay = today.Date.AddDays(1).AddTicks(-1);

                var startOfMonth = new DateTime(today.Year, today.Month, 1);
                var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);

                int casesCountToday = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfDay && r.Datetime <= endOfDay)
                     .Select(r => r.AccountNo) // Select the accountno
                    .Distinct() // Get distinct accountno values
                    .Count(); // Count the distinct values

                int casesCountThisMonth = eventTableQuery
                    .Where(r => r.Datetime != null && r.Datetime >= startOfMonth && r.Datetime <= endOfMonth)
                    .Select(r => r.AccountNo) // Select the accountno
                    .Distinct() // Get distinct accountno values
                    .Count(); // Count the distinct values




                var totalCasesList = allAgents.Select(UserName => recordDatasQuery.Count(r => r.Agent == UserName)).ToList(); /*recordDatasQuery.Where().Count();*/

                int totalCases = allAgents.Sum(UserName => recordDatasQuery.Count(r => r.Agent == UserName));


                int totalCasesNotWorked = recordDatasQuery
                    .Where(r => r.ModifiedDate == null)
                    .Count();

                int callbackCountThisMonth = recordDatasQuery
                    .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
                    .Count();


                List<string> allSegments = records
                .Where(r => !string.IsNullOrEmpty(r.Segments)) // Ensure we only consider non-null and non-empty segments
                .Select(r => r.Segments)
                .Distinct()
                .OrderBy(segment => segment)
                .ToList();


                List<int> segmentCounts = allSegments
                    .Select(segment => recordDatasQuery
                        .Count(r => r.Segments == segment))
                    .ToList();



                List<int> segmentNotWorked = allSegments
                    .Select(segment => recordDatasQuery
                        .Count(r => r.Segments == segment && r.ModifiedDate == null))
                    .ToList();

                var distinctSegments = recordDatasQuery
                    .Where(r => !string.IsNullOrEmpty(r.Segments))
                    .Select(r => r.Segments)
                    .OrderBy(segment => segment)
                    .Distinct()
                    .ToList();

                var segment2Counts = recordDatasQuery
                    .Where(r => !string.IsNullOrEmpty(r.Segments))
                    .GroupBy(r => r.DerbyBatch)
                    .OrderBy(g => g.Key)
                    .Select(g => g.Count())
                    .ToList();

                List<string> allbatches = records
                   .Where(r => !string.IsNullOrEmpty(r.DerbyBatch))
                   .Select(r => r.DerbyBatch)
                   .Distinct()
                   .OrderBy(DerbyBatch => DerbyBatch)
                   .ToList();

                List<int> batchWorked = allbatches
                .Select(batch => recordDatasQuery
                    .Where(r => !string.IsNullOrEmpty(r.DerbyBatch) && r.ModifiedDate != null && r.DerbyBatch == batch)
                    .Count())
                .ToList();

               

                List<int> totalBatches = allbatches
                   .Select(DerbyBatch => recordDatasQuery
                       .Count(r => r.DerbyBatch == DerbyBatch))
                   .ToList();

                List<int> batchNotWorked = allbatches
                .Select(batch => recordDatasQuery
                    .Where(r => !string.IsNullOrEmpty(r.DerbyBatch) && r.ModifiedDate == null && r.DerbyBatch == batch)
                    .Count())
                .ToList();

                var distinctSegments2 = recordDatasQuery
                    .Where(r => !string.IsNullOrEmpty(r.Segments))
                    .Select(r => r.DerbyBatch)
                    .OrderBy(DerbyBatch => DerbyBatch)
                    .Distinct()
                    .ToList();

                
                ViewBag.totalCasesList = totalCasesList;
                ViewBag.allAgents = fliterAllAgents;
                ViewBag.allbatches = allbatches;
                ViewBag.totalCasesNotWorked = totalCasesNotWorked;
                ViewBag.totalCases = totalCases;
                ViewBag.segmentNotWorked = segmentNotWorked;
                ViewBag.TotalBatches = totalBatches;
                ViewBag.BatchNotWorked = batchNotWorked;
                ViewBag.BatchWorked = batchWorked;
                ViewBag.segment2Counts = segment2Counts;
                ViewBag.distinctSegments2 = distinctSegments2;
                ViewBag.distinctSegments = distinctSegments;
                ViewBag.segmentCounts = segmentCounts;
                ViewBag.recentEvents = casesCountToday;
                ViewBag.callbackCountThisMonth = callbackCountThisMonth;
                ViewBag.CasesCountThisMonth = casesCountThisMonth;

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

                var agentCallBackCountToday = recordDatasQuery
                    .Where(r => r.CallbackTime.HasValue && r.CallbackTime >= startOfDay && r.CallbackTime <= endOfDay)
                    .GroupBy(r => r.ModifiedAgent)
                    .Select(g => new { ModifiedAgent = g.Key, CallBackCountToday = g.Count() })
                    .ToList();

                var agentCallBackCountPrev = recordDatasQuery
                    .Where(r => r.CallbackTime.HasValue && r.CallbackTime <= startOfDay && r.CallbackTime != new DateTime(2000, 1, 1))
                    .GroupBy(r => r.ModifiedAgent)
                    .Select(g => new { ModifiedAgent = g.Key, CallBackCountPrev = g.Count() })
                    .ToList();

                var agentCallBackCount = allAgents
                    .Select(agent => new AgentCasesViewModel
                    {

                        Agent = agent,
                        CallBackCountToday = agentCallBackCountToday.FirstOrDefault(ac => ac.ModifiedAgent == agent)?.CallBackCountToday ?? 0,
                        CallBackCountPrev = agentCallBackCountPrev.FirstOrDefault(ac => ac.ModifiedAgent == agent)?.CallBackCountPrev ?? 0
                    })
                    .ToList();

                ViewBag.AgentCallBackCount = agentCallBackCount;

                return View(agentsWithCasesCount);
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


            
                //db.EventTables.Find(int.Parse(form["Id"].ToString()));
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
            //if (!string.IsNullOrEmpty(form["CallbackTime"]))
            //{
            //recordData.CallbackTime = DateTime.Parse(form["CallbackTime"]);
            //}
            //else
            //{
            //// Assign a specific default value if CallbackTime is null or empty
            //recordData.CallbackTime = new DateTime(2000, 1, 1); // Example default value
            //}


            db.Entry(recordData).State = System.Data.Entity.EntityState.Modified;




            db.SaveChanges();
            return RedirectToAction("RealEditAllocation", new { id = form["Id"].ToString(), accountno = form["AccountNo"].ToString() });

            //return RedirectToAction("RealEditAllocation");
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

            if(drop2 == "Batch 1")
            {
                records = records.Where(et => et.DerbyBatch == "Batch 1").AsQueryable();
            }
            else if (drop2 == "Batch 2")
            {
                records = records.Where(et => et.DerbyBatch == "Batch 2").AsQueryable();
            }
            else if (drop2 == "Batch 3")
            {
                records = records.Where(et => et.DerbyBatch == "Batch 3").AsQueryable();
            }
            else if (drop2 == "Batch 4")
            {
                records = records.Where(et => et.DerbyBatch == "Batch 4").AsQueryable();
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

             if(drop4 == "Arshad")
            {
                records = records.Where(et => et.Agent == "Arshad").AsQueryable();
            }
            else if (drop4 == "Shoeb")
            {
                records = records.Where(et => et.Agent == "Shoeb").AsQueryable();
            }
            else if (drop4 == "Joshua")
            {
                records = records.Where(et => et.Agent == "Joshua").AsQueryable();
            }
              else if (drop4 == "TestAgent")
            {
                records = records.Where(et => et.Agent == "TestAgent").AsQueryable();
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
                .OrderBy(item => item.Id) // Ensure stable ordering
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

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

            IQueryable<RecordData> records = db.RecordDatas;
            IQueryable<EventTable> events = db.EventTables;

            DateTime today = DateTime.Today;
            var startOfDay = today.Date;
            var endOfDay = today.Date.AddDays(1).AddTicks(-1);
            var startOfMonth = new DateTime(today.Year, today.Month, 1);
            var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);

           

            if (state == "callback")
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
                    records = records.Where(item => item.Segments == query);
                }
                else if (section == "notworked")
                {
                    records = records.Where(item => item.Segments == query && item.ModifiedDate == null);
                }
                //ViewBag.Data = records;
                var settings = new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                };

                ViewBag.Data = JsonConvert.SerializeObject(records, settings);
                return PartialView("_AgentCases", records);

            }
            else if (state == "segment2")
            {
                if (section == "notworked")
                {
                    records = records.Where(item => item.DerbyBatch == query && item.ModifiedDate == null);
                }
                else if (section == "worked")
                {
                    records = records.Where(item => item.DerbyBatch == query && item.ModifiedDate != null);
                }
                //if (button == "clicked")
                //{
                //    return DownloadAgentCases(records.ToList());
                //}
                //ViewBag.Data = records;
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
                //if (button == "clicked")
                //{
                //    return DownloadAgentCases(records.ToList());
                //}
                //ViewBag.Data = records;
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
                    // Convert uploaded file to byte array
                    aspNetUser.Images = ConvertToBytes(file);

                    // Update existing user
                    db.Entry(aspNetUser).State = EntityState.Modified;
                    db.SaveChanges();
                }

                //user.Images = ConvertToBytes(file);


                //db.AspNetUsers.Add(user);
                //db.SaveChanges();
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
        //private void DownloadAgentCases(List<RecordData> records)
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("Agent Cases");
        //        var currentRow = 1;

        //        // Add column headers
        //        worksheet.Cell(currentRow, 1).Value = "Agent";
        //        worksheet.Cell(currentRow, 2).Value = "CallbackTime";
        //        worksheet.Cell(currentRow, 3).Value = "ModifiedDate";
        //        worksheet.Cell(currentRow, 4).Value = "Segments";
        //        worksheet.Cell(currentRow, 5).Value = "DerbyBatch";

        //        // Add data rows
        //        foreach (var record in records)
        //        {
        //            currentRow++;
        //            worksheet.Cell(currentRow, 1).Value = record.Agent;
        //            worksheet.Cell(currentRow, 2).Value = record.CallbackTime;
        //            worksheet.Cell(currentRow, 3).Value = record.ModifiedDate;
        //            worksheet.Cell(currentRow, 4).Value = record.Segments;
        //            worksheet.Cell(currentRow, 5).Value = record.DerbyBatch;
        //        }

        //        using (var stream = new MemoryStream())
        //        {
        //            workbook.SaveAs(stream);
        //            stream.Position = 0;

        //            var fileName = $"AgentCases_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
        //            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        //            Response.Clear();
        //            Response.ContentType = contentType;
        //            Response.Headers.Add("Content-Disposition", $"attachment; filename={fileName}");
        //            Response.OutputStream.Write(stream.ToArray(), 0, stream.ToArray().Length);
        //            Response.OutputStream.Flush();
        //        }
        //    }
        //}

        //public ActionResult DownloadAgentCases(List<RecordData> records)
        //{
        //    //List<RecordData> records2 = (List<RecordData>)Session["Records"];



        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //    using (var package = new ExcelPackage())
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Agent Cases");
        //        var currentRow = 1;

        //        // Add column headers
        //        worksheet.Cells[currentRow, 1].Value = "Agent";
        //        worksheet.Cells[currentRow, 2].Value = "CallbackTime";
        //        worksheet.Cells[currentRow, 3].Value = "ModifiedDate";
        //        worksheet.Cells[currentRow, 4].Value = "Segments";
        //        worksheet.Cells[currentRow, 5].Value = "DerbyBatch";

        //        // Add data rows
        //        foreach (var record in records)
        //        {
        //            currentRow++;
        //            worksheet.Cells[currentRow, 1].Value = record.Agent;
        //            worksheet.Cells[currentRow, 2].Value = record.CallbackTime;
        //            worksheet.Cells[currentRow, 3].Value = record.ModifiedDate;
        //            worksheet.Cells[currentRow, 4].Value = record.Segments;
        //            worksheet.Cells[currentRow, 5].Value = record.DerbyBatch;
        //        }

        //        using (var stream = new MemoryStream())
        //        {
        //            package.SaveAs(stream);
        //            stream.Position = 0;

        //            var fileName = $"AgentCases_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        //            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        //            //Response.Clear();
        //            //Response.ContentType = contentType;
        //            //Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
        //            //stream.CopyTo(Response.OutputStream);
        //            //Response.Flush();
        //            //Response.End();



        //            Response.Clear();
        //            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //            Response.AddHeader("content-disposition", "attachment;filename= " + fileName);
        //            Response.BinaryWrite(package.GetAsByteArray());
        //            Response.Flush();
        //            Response.SuppressContent = true;
        //            HttpContext.ApplicationInstance.CompleteRequest();
        //            string ipAddress = System.Web.HttpContext.Current.Request.UserHostAddress;
        //            return File(Response.OutputStream, Response.ContentType);
        //        }
        //    }


        //}


        //public ActionResult DownloadAgentCases(string records)
        //{
        //    List<RecordData> recordList = JsonConvert.DeserializeObject<List<RecordData>>(records);


        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    using (var package = new ExcelPackage())
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Agent Cases");
        //        var currentRow = 1;
        //        // Add column headers
        //        worksheet.Cells[currentRow, 1].Value = "Agent";
        //        worksheet.Cells[currentRow, 2].Value = "CallbackTime";
        //        worksheet.Cells[currentRow, 3].Value = "ModifiedDate";
        //        worksheet.Cells[currentRow, 4].Value = "Segments";
        //        worksheet.Cells[currentRow, 5].Value = "DerbyBatch";
        //        // Add data rows
        //        foreach (var record in recordList)
        //        {
        //            currentRow++;
        //            worksheet.Cells[currentRow, 1].Value = record.Agent;
        //            worksheet.Cells[currentRow, 2].Value = record.CallbackTime;
        //            worksheet.Cells[currentRow, 3].Value = record.ModifiedDate;
        //            worksheet.Cells[currentRow, 4].Value = record.Segments;
        //            worksheet.Cells[currentRow, 5].Value = record.DerbyBatch;
        //        }
        //        using (var stream = new MemoryStream())
        //        {
        //            package.SaveAs(stream);
        //            stream.Position = 0;

        //            var fileName = $"AgentCases_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        //            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        //            //Response.Clear();
        //            //Response.ContentType = contentType;
        //            //Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
        //            //stream.CopyTo(Response.OutputStream);
        //            //Response.Flush();
        //            //Response.End();



        //            Response.Clear();
        //            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //            Response.AddHeader("content-disposition", "attachment;filename= " + fileName);
        //            Response.BinaryWrite(package.GetAsByteArray());
        //            Response.Flush();
        //            Response.SuppressContent = true;
        //            HttpContext.ApplicationInstance.CompleteRequest();
        //            string ipAddress = System.Web.HttpContext.Current.Request.UserHostAddress;
        //            return File(Response.OutputStream, Response.ContentType);
        //        }
        //    }
        //}



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
                            worksheet.Cells[currentRow, i + 1].Value = properties[i].Name;
                        }

                        // Add data rows
                        foreach (var record in recordList)
                        {
                            currentRow++;
                            for (int i = 0; i < properties.Length; i++)
                            {
                                worksheet.Cells[currentRow, i + 1].Value = properties[i].GetValue(record)?.ToString();
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




        //[HttpPost]
        //public ActionResult _AgentCases(string query,string section, string state, string button)
        //{
        //    CPVDBEntities db = new CPVDBEntities();
        //    IQueryable<RecordData> records = db.RecordDatas;

        //    DateTime today = DateTime.Today;
        //    var startOfDay = today.Date;
        //    var endOfDay = today.Date.AddDays(1).AddTicks(-1);
        //    var startOfMonth = new DateTime(today.Year, today.Month, 1);
        //    var endOfMonth = startOfMonth.AddMonths(1).AddTicks(-1);

        //    if (state == "callback")
        //    {
        //        if(section == "day")
        //        {
        //            records = records.Where(item => item.Agent == query && item.CallbackTime >= startOfDay && item.CallbackTime <= endOfDay);

        //        }
        //        else if (section == "prev")
        //        {
        //            records = records.Where(item => item.Agent == query && item.CallbackTime >= startOfMonth && item.CallbackTime <= startOfDay);
        //        }

        //    }
        //    else if (state == "event")
        //    {
        //        if (section == "day")
        //        {
        //            records = records.Where(item => item.Agent == query && item.ModifiedDate >= startOfDay && item.ModifiedDate <= endOfDay);

        //        }
        //        else if (section == "month")
        //        {
        //            records = records.Where(item => item.Agent == query && item.ModifiedDate >= startOfMonth && item.ModifiedDate <= endOfMonth);
        //        }

        //    }
        //    else if (state == "segment1")
        //    {

        //        if (section == "all")
        //        {
        //            records = records.Where(item => item.Segments == query);

        //        }
        //        else if (section == "notworked")
        //        {
        //            records = records.Where(item => item.Segments == query && item.ModifiedDate == null);
        //        }

        //    } 
        //    else if (state == "segment2")
        //    {
        //        if (section == "notworked")
        //        {
        //            records = records.Where(item => item.DerbyBatch == query && item.ModifiedDate == null);

        //        }
        //        else if (section == "worked")
        //        {
        //            records = records.Where(item => item.DerbyBatch == query && item.ModifiedDate != null);
        //        }
        //    }   
        //    else
        //    {
        //        records = records.Where(item => item.Agent == query);
        //    }
        //    if (button == "clicked")
        //    {
        //         DownloadAgentCases(records.ToList());
        //        return PartialView("_AgentCases", records);
        //    }
        //    else
        //    {
        //        return PartialView("_AgentCases", records);
        //    }


        //}

        //public ActionResult DownloadAgentCases(List<RecordData> records)
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("AgentCases");
        //        var currentRow = 1;

        //        // Add header row
        //        worksheet.Cell(currentRow, 1).Value = "Agent";
        //        worksheet.Cell(currentRow, 2).Value = "CallbackTime";
        //        worksheet.Cell(currentRow, 3).Value = "ModifiedDate";
        //        worksheet.Cell(currentRow, 4).Value = "Segments";
        //        worksheet.Cell(currentRow, 5).Value = "DerbyBatch";

        //        // Add data rows
        //        foreach (var record in records)
        //        {
        //            currentRow++;
        //            worksheet.Cell(currentRow, 1).Value = record.Agent;
        //            worksheet.Cell(currentRow, 2).Value = record.CallbackTime;
        //            worksheet.Cell(currentRow, 3).Value = record.ModifiedDate;
        //            worksheet.Cell(currentRow, 4).Value = record.Segments;
        //            worksheet.Cell(currentRow, 5).Value = record.DerbyBatch;
        //        }

        //        using (var stream = new MemoryStream())
        //        {
        //            workbook.SaveAs(stream);
        //            stream.Position = 0; // Reset the stream position to the beginning

        //            var content = stream.ToArray();

        //            return new FileContentResult(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //            {
        //                FileDownloadName = "AgentCases.xlsx"
        //            };
        //        }
        //    }
        //}



        //[HttpPost]
        //public ActionResult _FilterRecords(string query)
        //{
        //    CPVDBEntities db = new CPVDBEntities();

        //    var results = db.RecordDatas.ToList();

        //    return PartialView("_Records", results);
        //}






        //[HttpPost]
        //public ActionResult _Records(string query)
        //{

        //    CPVDBEntities db = new CPVDBEntities();
        //    RecordData recordData = new RecordData();
        //    EventTable eventTable = new EventTable();
        //    var DateTime = db.EventTables.ToList();
        //    var results1 = db.RecordDatas.Take(200).ToList();
        //    var userName = User.Identity.GetUserName();

        //    if (query == "")
        //    {
        //        if (User.IsInRole("Agent"))
        //        {
        //            results1 = db.RecordDatas
        //               .Where(item => item.Agent == userName)
        //               .ToList();
        //        }

        //    }
        //    else if (query != "")
        //    {
        //            results1 = db.RecordDatas
        //            .Where(item => item.CustomerName == query ||
        //                           item.AccountNo == query)
        //            .ToList();

        //    }
        //        else
        //        {
        //            results1 = db.RecordDatas.ToList();
        //        }


        //    return PartialView("_Records", results1);
        //}

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
                                IPTelephone_Billing= worksheet1.Cells[row, 19].Value?.ToString(),
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


                                
                            };

                            if (!(caseEntity1.OS_Billing == null || caseEntity1.OS_Billing == "-" || caseEntity1.OS_Billing.Trim() == "0") && !(caseEntity1.ExpectedRenewalFee == null || caseEntity1.ExpectedRenewalFee == "-" || caseEntity1.ExpectedRenewalFee.Trim() == "0"))
                            {
                                caseEntity1.Segments = "Bounced Cheque and Renewal";
                            }
                            else if(!(caseEntity1.OS_Billing == null || caseEntity1.OS_Billing == "-" || caseEntity1.OS_Billing.Trim() == "0"))
                            {
                                caseEntity1.Segments = "Bounced Cheque";
                            }
                            else if(!(caseEntity1.ExpectedRenewalFee == null || caseEntity1.ExpectedRenewalFee == "-" || caseEntity1.ExpectedRenewalFee.Trim() == "0"))
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



                            // Insert mobile numbers into MobileNos table



                        }

                       


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

   

        void InsertMobileNo(CPVDBEntities context, int recordId, string mobileNo)
        {
            if (!string.IsNullOrEmpty(mobileNo))
            {
                var mobileRecord = new MobileNo
                {
                    RecordId = recordId,
                    Numbers = mobileNo
                };

                context.Set<MobileNo>().Add(mobileRecord);
                context.SaveChanges();
            }
        }

        void InsertEmailId(CPVDBEntities context, int recordId, string emailId)
        {
            if (!string.IsNullOrEmpty(emailId))
            {
                var emailRecord = new EmailId
                {
                    RecordId = recordId,
                    Emails = emailId
                };

                context.Set<EmailId>().Add(emailRecord);
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

        public ActionResult DownloadCases(string Users, string CallType,string Disposition,string SubDisposition, DateTime? CallbackTime = null, DateTime? specificDate = null, DateTime? endDate = null)
        {
            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = db.EventTables.AsQueryable();
            var caseTables = query.ToList();

            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            List<EventTable> eventTables = new List<EventTable>();
            List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            bouncedRecords = db.BouncedRecords.ToList();    

            var key = "";

            if (Users != "")
            {
                query = query.Where(et => et.Agent == Users);
                var hasRows = query.Any();
                caseTables = query.ToList();

            }

            if (CallType != "-")
            {
                query = query.Where(et => et.CallType == CallType);
                caseTables = query.ToList();
            }

            if (Disposition != "-")
            {
                query = query.Where(et => et.Dispo == Disposition);
                caseTables = query.ToList();
            }

            if (SubDisposition != "-")
            {
                query = query.Where(et => et.SubDispo == SubDisposition);
                caseTables = query.ToList();
            }

            if (CallbackTime != null)
            {
                query = query.Where(et => et.CallbackTime == CallbackTime);
                caseTables = query.ToList();
            }

            if (specificDate != null)
            {
                query = query.Where(et => et.Datetime >= specificDate);
                caseTables = query.ToList();
            }

            if (endDate != null)
            {
                query = query.Where(et => et.Datetime <= endDate);
                caseTables = query.ToList();
            }





            var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing",
                "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
                "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee",
                 "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
                "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            var header1 = new List<string>()
            {
                "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments", "Segments"
            };

            var header2 = new List<string>()
            {
                "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments", "Segments"
            };

            var bouncedDetailHeaders = new List<string>()
            {
                 "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", "ChequeDate","TotalAmount"
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

                //foreach (var items in recordDatas)
                //{


                //    //recordDatas = db.RecordDatas.Where(rd => rd.AccountNo != key).ToList();

                //    recordDatas = db.RecordDatas
                //    .Where(rd => !accountNoList.Contains(rd.AccountNo))
                //    .ToList();

                //    foreach (var itemcase in recordDatas)
                //    {
                //        col = 1;
                //        foreach (var headName in header)
                //        {
                //            var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

                //            object value = property.GetValue(itemcase);

                //            if (property.PropertyType == typeof(DateTime))
                //            {
                //                worksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                //            }
                //            else
                //            {
                //                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
                //            }


                //        }
                //        for (int i = 0; i < 10; i++)
                //        {
                //            worksheet.Cells[row, col++].Value = "-";
                //        }
                //        row++;
                //    }



                //}

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
                Response.AddHeader("content-disposition", "attachment;filename=Caselist" + (specificDate.HasValue ? specificDate.Value.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd")) + ".xlsx");

                Response.BinaryWrite(package.GetAsByteArray());
                Response.Flush();
                Response.SuppressContent = true;
                HttpContext.ApplicationInstance.CompleteRequest();

                return File(Response.OutputStream, Response.ContentType);
            }
        }

        public ActionResult DownloadRecordCases(string Users, string CallType, string Disposition, string SubDisposition, DateTime? CallbackTime = null, DateTime? specificDate = null, DateTime? endDate = null)
        {
            //CPVDBEntities db = new CPVDBEntities();
            //HttpResponseBase Response = HttpContext.Response;
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            //List<EventTable> eventTables = new List<EventTable>();
            //List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            //bouncedRecords = db.BouncedRecords.ToList();    

            //var key = "";
            //eventTables =  db.EventTables
            //                   .ToList();



            //if (specificDate.HasValue && endDate.HasValue)
            //{
            //    eventTables = db.EventTables
            //                    .Where(et => et.Datetime >= specificDate.Value && et.Datetime <= endDate.Value)
            //                    .ToList();
            //}
            //else if (specificDate.HasValue)
            //{
            //    eventTables = db.EventTables
            //                    .Where(et => et.Datetime > specificDate.Value)
            //                    .ToList();


            //}

            //var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing", 
            //    "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
            //    "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee", 
            //     "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
            //    "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            //var header1 = new List<string>()
            //{
            //    "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments", "Segments"
            //};

            // var header2 = new List<string>()
            //{
            //    "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments", "Segments"
            //};

            //var bouncedDetailHeaders = new List<string>()
            //{
            //     "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", "ChequeDate","TotalAmount"
            //};


            //using (var package = new ExcelPackage())
            //{
            //    int col = 1;
            //    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            //    var combinedList = header.Concat(header2).ToList();
            //    List<string> accountNoList = new List<string>();

            //    foreach (var headerName in combinedList )
            //    {
            //        worksheet.Cells[1, col++].Value = headerName.ToString();
            //    }



            //    int row = 2;

            //    foreach (var items in eventTables)
            //    {
            //        col = 31;
            //        foreach (var headName in header1)
            //        {
            //            var property = typeof(EventTable).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);
            //            object value = property.GetValue(items);

            //            if (property != null && property.Name == "AccountNo")
            //            {
            //                key = value.ToString();
            //                accountNoList.Add(value?.ToString());
            //            }

            //            if (value != null && value.ToString() == "1/1/2000 12:00:00 AM")
            //            {

            //                worksheet.Cells[row, col++].Value = "-";
            //            }
            //            else
            //            {
            //                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
            //            }
            //        }

            //        recordDatas = db.RecordDatas.Where(rd => rd.AccountNo == key).ToList();

            //        foreach (var itemcase in recordDatas)
            //        {

            //            col = 1;
            //            foreach (var headName in header)
            //            {
            //                var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

            //                object value = property.GetValue(itemcase);

            //                if (property.PropertyType == typeof(DateTime))
            //                {
            //                    worksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
            //                }
            //                else
            //                {
            //                worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
            //                }
            //            }

            //        }

            //        row++;
            //    }

            //    foreach (var items in recordDatas)
            //    {


            //        //recordDatas = db.RecordDatas.Where(rd => rd.AccountNo != key).ToList();

            //        recordDatas = db.RecordDatas
            //        .Where(rd => !accountNoList.Contains(rd.AccountNo))
            //        .ToList();

            //        foreach (var itemcase in recordDatas)
            //        {
            //            col = 1;
            //            foreach (var headName in header)
            //            {
            //                var property = typeof(RecordData).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

            //                object value = property.GetValue(itemcase);

            //                if (property.PropertyType == typeof(DateTime))
            //                {
            //                    worksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
            //                }
            //                else
            //                {
            //                    worksheet.Cells[row, col++].Value = !string.IsNullOrEmpty(value?.ToString()) ? value.ToString() : "-";
            //                }


            //            }
            //            for (int i = 0; i < 10; i++)
            //            {
            //                worksheet.Cells[row, col++].Value = "-";
            //            }
            //            row++;
            //        }



            //    }

            //    var bouncedWorksheet = package.Workbook.Worksheets.Add("BouncedDetails");
            //    col = 1;
            //    foreach (var headerName in bouncedDetailHeaders)
            //    {
            //        bouncedWorksheet.Cells[1, col++].Value = headerName.ToString();
            //    }


            //    row = 2;

            //    foreach (var bouncedDetail in bouncedRecords)
            //    {
            //        col = 1;
            //        foreach (var headerName in bouncedDetailHeaders)
            //        {
            //            var property = typeof(BouncedRecord).GetProperty(headerName, BindingFlags.Public | BindingFlags.Instance);
            //            object value = property.GetValue(bouncedDetail);

            //            if (property.PropertyType == typeof(DateTime))
            //            {
            //                bouncedWorksheet.Cells[row, col++].Value = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
            //            }
            //            else
            //            {
            //                bouncedWorksheet.Cells[row, col++].Value = value != null ? value.ToString() : string.Empty;
            //            }
            //        }


            //        row++;
            //    }

            //    Response.Clear();
            //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //    Response.AddHeader("content-disposition", "attachment;filename=Caselist" + (specificDate.HasValue ? specificDate.Value.ToString("yyyy-MM-dd") : DateTime.Today.ToString("yyyy-MM-dd")) + ".xlsx");

            //    Response.BinaryWrite(package.GetAsByteArray());
            //    Response.Flush();
            //    Response.SuppressContent = true;
            //    HttpContext.ApplicationInstance.CompleteRequest();

            //    return File(Response.OutputStream, Response.ContentType);
            //}


            CPVDBEntities db = new CPVDBEntities();
            HttpResponseBase Response = HttpContext.Response;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var query = db.RecordDatas.AsQueryable();
            var caseTables = query.ToList();

            List<RecordData> recordDatas = new List<RecordData>(); // Initialize with an empty list
            List<EventTable> eventTables = new List<EventTable>();
            List<BouncedRecord> bouncedRecords = new List<BouncedRecord>();

            bouncedRecords = db.BouncedRecords.ToList();

            var key = "";

            if (Users != "")
            {
                query = query.Where(et => et.Agent == Users);
                var hasRows = query.Any();
                caseTables = query.ToList();

            }

            if (CallType != "-")
            {
                query = query.Where(et => et.CallType == CallType);
                caseTables = query.ToList();
            }

            if (Disposition != "-")
            {
                query = query.Where(et => et.Disposition == Disposition);
                caseTables = query.ToList();
            }

            if (SubDisposition != "-")
            {
                query = query.Where(et => et.SubDisposition == SubDisposition);
                caseTables = query.ToList();
            }

            if (CallbackTime != null)
            {
                query = query.Where(et => et.CallbackTime == CallbackTime);
                caseTables = query.ToList();
            }

            if (specificDate != null)
            {
                query = query.Where(et => et.ModifiedDate >= specificDate);
                caseTables = query.ToList();
            }

            if (endDate != null)
            {
                query = query.Where(et => et.ModifiedDate <= endDate);
                caseTables = query.ToList();
            }





            var header = new List<string>() {  "AccountNo","Agent","CustomerName", "BCheque", "BCheque_P", "IPTelephone_Billing",
                "Utility_Billing", "Others", "OS_Billing", "License_expiry", "Contact_Person","ModifiedDate", "Nationality", "Mobile1",
                "Mobile2", "Mobile3", "Mobile4", "Email_1", "Email_2", "Email_3", "ExpectedRenewalFee",
                 "SRNumber", "EmployeeVisaQuota", "EmployeeVisaUtilized", "ProjectBundleName",
                "LicenseType", "FacilityType", "NoYears", "DerbyBatch", "CallType"};


            var header1 = new List<string>()
            {
                "AccountNo","Agent", "DialedNumber", "EmailUsed", "Dispo", "SubDispo", "CallbackTime", "Comments", "Segments"
            };

            var header2 = new List<string>()
            {
                "Account No","Event Created by", "Dialed Number", "Email Used", "Disposition", "Sub Disposition", "CallbackTime", "Comments", "Segments"
            };

            var bouncedDetailHeaders = new List<string>()
            {
                 "AccountNo","ChequeNumber", "ReasonCode","Text","DateBounced", "ChequeDate","TotalAmount"
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

                    eventTables = db.EventTables
                                     .Where(rd => rd.AccountNo == key)
                                     .GroupBy(rd => rd.AccountNo)
                                     .Select(g => g.OrderByDescending(rd => rd.Datetime).FirstOrDefault())
                                     .ToList();

                    foreach (var itemcase in eventTables)
                    {

                        col = 31;
                        foreach (var headName in header1)
                        {
                            var property = typeof(EventTable).GetProperty(headName, BindingFlags.Public | BindingFlags.Instance);

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