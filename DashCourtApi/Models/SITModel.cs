namespace DashCourtApi.Models
{
    public class SITModel
    {
        public int Year { get; set; }
        public decimal DelaysPercentage { get; set; }
        public int Rejected { get; set; }
        public int Hearings { get; set; }
        public string CaseType { get; set; }
        public string Procedure { get; set; }
        public string Court { get; set; }
        public string District { get; set; }
        public string OriginalCycle { get; set; }
    }
}
