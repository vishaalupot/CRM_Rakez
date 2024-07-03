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
    
    public partial class RecordData
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RecordData()
        {
            this.EmailIds = new HashSet<EmailId>();
            this.MobileNos = new HashSet<MobileNo>();
        }
    
        public int Id { get; set; }
        public string AccountNo { get; set; }
        public string CustomerName { get; set; }
        public string BCheque { get; set; }
        public string BCheque_P { get; set; }
        public string IPTelephone_Billing { get; set; }
        public string Utility_Billing { get; set; }
        public string Others { get; set; }
        public string OS_Billing { get; set; }
        public string License_expiry { get; set; }
        public string Contact_Person { get; set; }
        public string Nationality { get; set; }
        public string Mobile1 { get; set; }
        public string Mobile2 { get; set; }
        public string Mobile3 { get; set; }
        public string Mobile4 { get; set; }
        public string Email_1 { get; set; }
        public string Email_2 { get; set; }
        public string Email_3 { get; set; }
        public string Disposition { get; set; }
        public string SubDisposition { get; set; }
        public string Comments { get; set; }
        public string ChangeStatus { get; set; }
        public Nullable<System.DateTime> CallbackTime { get; set; }
        public string Agent { get; set; }
        public Nullable<System.DateTime> ModifiedDate { get; set; }
        public string CloseAccount { get; set; }
        public string DormantAccount { get; set; }
        public string InsufficientFunds { get; set; }
        public string OtherReason { get; set; }
        public string SignatureIrregular { get; set; }
        public string TechnicalReason { get; set; }
        public string BOthers { get; set; }
        public string Segments { get; set; }
        public string DialedNumber { get; set; }
        public string TenacyFacilityType { get; set; }
        public string ExpectedRenewalFee { get; set; }
        public string SRNumber { get; set; }
        public string DeRegFee { get; set; }
        public string EmployeeVisaQuota { get; set; }
        public string EmployeeVisaUtilized { get; set; }
        public string ProjectBundleName { get; set; }
        public string LicenseType { get; set; }
        public string FacilityType { get; set; }
        public string NoYears { get; set; }
        public string AgentsName { get; set; }
        public string DispositionSecond { get; set; }
        public string EmailUsed { get; set; }

        public string DerbyBatch { get; set; }
        public string CallType { get; set; }
        public string ModifiedAgent { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EmailId> EmailIds { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MobileNo> MobileNos { get; set; }
    }
}
