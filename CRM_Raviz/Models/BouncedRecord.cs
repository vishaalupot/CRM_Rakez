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
    
    public partial class BouncedRecord
    {
        public int Id { get; set; }
        public string AccountNo { get; set; }
        public string ChequeNumber { get; set; }
        public string ReasonCode { get; set; }
        public string Text { get; set; }
        public Nullable<System.DateTime> DateBounced { get; set; }
        public Nullable<System.DateTime> ChequeDate { get; set; }
        public string TotalAmount { get; set; }
    }
}