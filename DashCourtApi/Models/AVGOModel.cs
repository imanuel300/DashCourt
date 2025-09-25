namespace DashCourtApi.Models
{
    public class AVGOModel
    {
        public int Year { get; set; }
        public string DateCycle { get; set; }
        public decimal AverageDays { get; set; }
        public int NumberOfCases { get; set; }
        public string CaseType { get; set; }
        public string Procedure { get; set; }
        public string Court { get; set; }
        public string District { get; set; }
        public string OriginalCycle { get; set; }
    }
}
