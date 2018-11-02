using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EC_to_SAP_Analysis_Console
{
    public class iCIMS
    {
        public string FullName { get; set; }
        public string StartDate { get; set; }
        public string PositionNumber { get; set; }
        public string RequistionID { get; set; }
        public string Country { get; set; }
        public string LastPreOnboard { get; set; }
        public string EventReason { get; set; }
    }

    public class ONB
    {
        public string EmployeeLogin { get; set; }
        public string PostHireVerificationStepStartdate { get; set; }
        public string PostHireVerificationStepEnddate { get; set; }
        public string GlobalPreOnboardingStartdate { get; set; }
        public string GlobalPreOnboardingEnddate { get; set; }
    }

    public class SAPNameValidation
    {
        public string PersonnelNumber { get; set; }
        public string EmployeeName { get; set; }
    }

    public class SAPHireCompletionDateData
    {
        public string PersonnelNumber { get; set; }
        public string HireCompletedDateinSAP { get; set; }
    }

    public class AliasCreation
    {
        public string PersonnelNumber { get; set; }
        public string AliasCreationDate { get; set; }
    }

    public class EC_TO_SAP
    {
        [DisplayName("Particulars")]
        public string Particulars { get; set; }
        [DisplayName("Full Name")]
        public string FullName { get; set; }
        [DisplayName("Hire Date")]
        public string StartDate { get; set; }
        [DisplayName(" Position Number")]
        public string PositionNumber { get; set; }
        [DisplayName("Requisition ID")]
        public string RequistionID { get; set; }
        [DisplayName("Country")]
        public string Country { get; set; }
        [DisplayName("Last Pre-Onboard: HCDT Completed date")]
        public string LastPreOnboard { get; set; }
        [DisplayName("Event Reason")]
        public string EventReason { get; set; }
        [DisplayName("Personnel number")]
        public string PersonnelNumber { get; set; }
        [DisplayName("Name Match")]
        public string EmployeeName { get; set; }
        [DisplayName("Post Hire Verification Step start Date")]
        public string PostHireVerificationStepStartdate { get; set; }
        [DisplayName("Post Hire Verification Step End Date")]
        public string PostHireVerificationStepEnddate { get; set; }
        [DisplayName("Global Pre-Onboarding start Date")]
        public string GlobalPreOnboardingStartdate { get; set; }
        [DisplayName("Global Pre-Onboarding End Date")]
        public string GlobalPreOnboardingEnddate { get; set; }
        [DisplayName("MPH Completed Date")]
        public string MPHCompletedDate { get; set; }
        [DisplayName("Hire Completed Date in SAP")]
        public string HireCompletedDateinSAP { get; set; }
        [DisplayName("Alias creation Date")]
        public string AliasCreationDate { get; set; }
    }
}
