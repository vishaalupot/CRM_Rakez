﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CRM_Raviz.Models
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class CPVDBEntities : DbContext
    {
        public CPVDBEntities()
            : base("name=CPVDBEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<AspNetRole> AspNetRoles { get; set; }
        public virtual DbSet<AspNetUserClaim> AspNetUserClaims { get; set; }
        public virtual DbSet<AspNetUserLogin> AspNetUserLogins { get; set; }
        public virtual DbSet<AspNetUser> AspNetUsers { get; set; }
        public virtual DbSet<RecordData> RecordDatas { get; set; }
        public virtual DbSet<BouncedRecord> BouncedRecords { get; set; }
        public virtual DbSet<EventTable> EventTables { get; set; }
        public virtual DbSet<Test_4> Test_4 { get; set; }
        public virtual DbSet<EmailId> EmailIds { get; set; }
        public virtual DbSet<MobileNo> MobileNos { get; set; }
        public virtual DbSet<AllocationData> AllocationDatas { get; set; }
    }
}
