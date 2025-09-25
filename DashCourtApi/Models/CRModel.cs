namespace DashCourtApi.Models
{
    public class CRModel
    {
        public int Inventory { get; set; }
        public int Year { get; set; }
        public decimal CR { get; set; }
        public int Opened { get; set; }
        public int Closed { get; set; }
        public string? CaseType { get; set; }
        public string? Procedure { get; set; }
        public string? Court { get; set; }
        public string? District { get; set; }
        public string? Cycle { get; set; }
        public string? OriginalCycle { get; set; }
    }
}
