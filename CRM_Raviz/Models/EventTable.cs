//------------------------------------------------------------------------------
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
    using System.Collections.Generic;
    
    public partial class EventTable
    {
        public int Id { get; set; }
        public string AccountNo { get; set; }
        public string CustomerName { get; set; }
        public Nullable<System.DateTime> Date { get; set; }
        public Nullable<System.TimeSpan> Time { get; set; }
        public string Dispo { get; set; }
        public string SubDispo { get; set; }
        public string Comments { get; set; }
        public string ChangeStatus { get; set; }
        public Nullable<System.DateTime> CallbackTime { get; set; }
        public Nullable<int> Record_Id { get; set; }
        public string Agent { get; set; }
        public Nullable<System.DateTime> ModifiedDate { get; set; }
        public string Segments { get; set; }
    }
}
