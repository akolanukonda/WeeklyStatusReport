using System.ComponentModel.DataAnnotations;

namespace WeeklyStatusReport.Models
{
    public class DigitalSupportTeam
    {
        // Sharepoint
        [Required]
        public int Sharepoint_Assigned { get; set; }

        [Required]
        public int Sharepoint_Closed { get; set; }

        [Required]
        public int Sharepoint_CarryForward { get; set; }

        // Digital- My Resource\Dragonboat
        [Required]
        public int Digital_MyResource_Assigned { get; set; }

        [Required]
        public int Digital_MyResource_Closed { get; set; }

        [Required]
        public int Digital_MyResource_CarryForward { get; set; }

        // Digital- Dot.com\E-Commerce
        [Required]
        public int Digital_Dotcom_Assigned { get; set; }

        [Required]
        public int Digital_Dotcom_Closed { get; set; }

        [Required]
        public int Digital_Dotcom_CarryForward { get; set; }

        // Compass
        [Required]
        public int Compass_Assigned { get; set; }

        [Required]
        public int Compass_Closed { get; set; }

        [Required]
        public int Compass_CarryForward { get; set; }

        // Doc Locator
        [Required]
        public int DocLocator_Assigned { get; set; }

        [Required]
        public int DocLocator_Closed { get; set; }

        [Required]
        public int DocLocator_CarryForward { get; set; }

        // CFirst\IDS
        [Required]
        public int CFirst_IDS_Assigned { get; set; }

        [Required]
        public int CFirst_IDS_Closed { get; set; }

        [Required]
        public int CFirst_IDS_CarryForward { get; set; }

        // NA Portal
        [Required]
        public int NAPortal_Assigned { get; set; }

        [Required]
        public int NAPortal_Closed { get; set; }

        [Required]
        public int NAPortal_CarryForward { get; set; }

        // Microsites\Others
        [Required]
        public int Microsites_Others_Assigned { get; set; }

        [Required]
        public int Microsites_Others_Closed { get; set; }

        [Required]
        public int Microsites_Others_CarryForward { get; set; }

        // ACN
        [Required]
        public int ACN_Assigned { get; set; }

        [Required]
        public int ACN_Closed { get; set; }

        [Required]
        public int ACN_CarryForward { get; set; }

        // Adhoc
        [Required]
        public int Adhoc_Assigned { get; set; }

        [Required]
        public int Adhoc_Closed { get; set; }

        [Required]
        public int Adhoc_CarryForward { get; set; }

        // Total Tickets (read-only)
        public int TotalTickets_Assigned { get; set; }

        public int TotalTickets_Closed { get; set; }

        public int TotalTickets_CarryForward { get; set; }

        [Required]
        public int Urgent { get; set; }

        // High Priority Tickets (read-only)
        [Required]
        public int HighPriorityTickets { get; set; }
    }
}
