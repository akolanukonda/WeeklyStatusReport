using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace WeeklyStatusReport.Models
{
    public class DigitalTeam
    {
        public string TeamName { get; set; }

        [Required(ErrorMessage = "Please select a team.")]
        public string SubTeamName {  get; set; }
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
