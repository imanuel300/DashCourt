namespace DashCourtApi.Models
{
    public class Inv3Model
    {
        public int BaseYear { get; set; }
        public int DaysInSystem { get; set; }
        public decimal Average { get; set; }
        public int CurrentInventory { get; set; }
        public int Year { get; set; }
        public string? YearAndMonth { get; set; }
        public string? CaseType { get; set; }
        public string? Procedure { get; set; }
        public string? Court { get; set; }
        public string? District { get; set; }
        public string? OriginalCycle { get; set; }
    }
}
