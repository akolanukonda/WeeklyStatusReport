using System.Web.Mvc;

namespace WeeklyStatusReport.Models
{
    public class CloudSailorsTeam
    {
        public string TeamName { get; set; }
        public string Week { get; set; }

        [AllowHtml]
        public required string Description { get; set; }
        public string Status { get; set; }
        public string Risks { get; set; }
        public string Accomplishments { get; set; }
        public string ClosureDate { get; set; }
        public string ProjectType { get; set; }
    }
}
