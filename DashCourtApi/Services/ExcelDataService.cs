using OfficeOpenXml;
using DashCourtApi.Models;

namespace DashCourtApi.Services
{
    public class ExcelDataService
    {
        private readonly string _excelFilesPath;
        private readonly bool _useMockData;

        public ExcelDataService(string excelFilesPath, bool useMockData)
        {
            _excelFilesPath = excelFilesPath;
            _useMockData = useMockData;
            ExcelPackage.License.SetNonCommercialOrganization("DashCourt");
        }

        public List<CRModel> GetCRData(string fileName = "CR.xlsx")
        {
            if (_useMockData)
            {
                return new List<CRModel>
                {
                    new CRModel { Inventory = 100, Year = 2023, CR = 0.75m, Opened = 50, Closed = 30, CaseType = "MockCR", Procedure = "MockProc", Court = "MockCourt", District = "MockDist", Cycle = "MockCycle", OriginalCycle = "MockOrigCycle" },
                    new CRModel { Inventory = 120, Year = 2023, CR = 0.80m, Opened = 60, Closed = 45, CaseType = "MockCR2", Procedure = "MockProc2", Court = "MockCourt2", District = "MockDist2", Cycle = "MockCycle2", OriginalCycle = "MockOrigCycle2" }
                };
            }

            var data = new List<CRModel>();
            var filePath = Path.Combine(_excelFilesPath, fileName);

            if (!File.Exists(filePath))
            {
                return data;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    data.Add(new CRModel
                    {
                        Inventory = GetCellValueAsInt(worksheet.Cells[row, 2]),
                        Year = GetCellValueAsInt(worksheet.Cells[row, 3]),
                        CR = GetCellValueAsDecimal(worksheet.Cells[row, 4], true), // true for percentage
                        Opened = GetCellValueAsInt(worksheet.Cells[row, 5]),
                        Closed = GetCellValueAsInt(worksheet.Cells[row, 6]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 7]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 8]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 9]),
                        District = GetCellValueAsString(worksheet.Cells[row, 10]),
                        Cycle = GetCellValueAsString(worksheet.Cells[row, 11]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 12])
                    });
                }
            }
            return data;
        }

        public List<AVGOModel> GetAVGOData(string fileName = "AVGO.xlsx")
        {
            if (_useMockData)
            {
                return new List<AVGOModel>
                {
                    new AVGOModel { Year = 2023, DateCycle = "Q1", AverageDays = 150.5m, NumberOfCases = 500, CaseType = "MockAVGO", Procedure = "MockProc", Court = "MockCourt", District = "MockDist", OriginalCycle = "MockOrigCycle" },
                    new AVGOModel { Year = 2023, DateCycle = "Q2", AverageDays = 140.0m, NumberOfCases = 550, CaseType = "MockAVGO2", Procedure = "MockProc2", Court = "MockCourt2", District = "MockDist2", OriginalCycle = "MockOrigCycle2" }
                };
            }

            var data = new List<AVGOModel>();
            var filePath = Path.Combine(_excelFilesPath, fileName);

            if (!File.Exists(filePath))
            {
                return data;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    data.Add(new AVGOModel
                    {
                        Year = GetCellValueAsInt(worksheet.Cells[row, 2]),
                        DateCycle = GetCellValueAsString(worksheet.Cells[row, 3]),
                        AverageDays = GetCellValueAsDecimal(worksheet.Cells[row, 4]),
                        NumberOfCases = GetCellValueAsInt(worksheet.Cells[row, 5]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 8]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 9]),
                        District = GetCellValueAsString(worksheet.Cells[row, 10]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 11])
                    });
                }
            }
            return data;
        }

        public List<SITModel> GetSITData(string fileName = "SIT.xlsx")
        {
            if (_useMockData)
            {
                return new List<SITModel>
                {
                    new SITModel { Year = 2023, DelaysPercentage = 0.20m, Rejected = 10, Hearings = 50, CaseType = "MockSIT", Procedure = "MockProc", Court = "MockCourt", District = "MockDist", OriginalCycle = "MockOrigCycle" },
                    new SITModel { Year = 2023, DelaysPercentage = 0.15m, Rejected = 8, Hearings = 45, CaseType = "MockSIT2", Procedure = "MockProc2", Court = "MockCourt2", District = "MockDist2", OriginalCycle = "MockOrigCycle2" }
                };
            }

            var data = new List<SITModel>();
            var filePath = Path.Combine(_excelFilesPath, fileName);

            if (!File.Exists(filePath))
            {
                return data;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    data.Add(new SITModel
                    {
                        Year = GetCellValueAsInt(worksheet.Cells[row, 2]),
                        DelaysPercentage = GetCellValueAsDecimal(worksheet.Cells[row, 3], true),
                        Rejected = GetCellValueAsInt(worksheet.Cells[row, 4]),
                        Hearings = GetCellValueAsInt(worksheet.Cells[row, 5]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 7]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 8]),
                        District = GetCellValueAsString(worksheet.Cells[row, 9]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 10])
                    });
                }
            }
            return data;
        }

        public List<Inv3Model> GetInv3Data(string fileName = "Inv3.xlsx")
        {
            if (_useMockData)
            {
                return new List<Inv3Model>
                {
                    new Inv3Model { BaseYear = 2022, DaysInSystem = 300, Average = 250.5m, CurrentInventory = 1000, Year = 2023, YearAndMonth = "2023-01", CaseType = "MockInv3", Procedure = "MockProc", Court = "MockCourt", District = "MockDist", OriginalCycle = "MockOrigCycle" },
                    new Inv3Model { BaseYear = 2022, DaysInSystem = 320, Average = 260.0m, CurrentInventory = 1100, Year = 2023, YearAndMonth = "2023-02", CaseType = "MockInv32", Procedure = "MockProc2", Court = "MockCourt2", District = "MockDist2", OriginalCycle = "MockOrigCycle2" }
                };
            }

            var data = new List<Inv3Model>();
            var filePath = Path.Combine(_excelFilesPath, fileName);

            if (!File.Exists(filePath))
            {
                return data;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    data.Add(new Inv3Model
                    {
                        BaseYear = GetCellValueAsInt(worksheet.Cells[row, 2]),
                        DaysInSystem = GetCellValueAsInt(worksheet.Cells[row, 3]),
                        Average = GetCellValueAsDecimal(worksheet.Cells[row, 4]),
                        CurrentInventory = GetCellValueAsInt(worksheet.Cells[row, 5]),
                        Year = GetCellValueAsInt(worksheet.Cells[row, 6]),
                        YearAndMonth = GetCellValueAsString(worksheet.Cells[row, 7]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 8]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 9]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 10]),
                        District = GetCellValueAsString(worksheet.Cells[row, 11]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 12])
                    });
                }
            }
            return data;
        }

        private int GetCellValueAsInt(ExcelRange cell)
        {
            return int.TryParse(cell?.Value?.ToString(), out int value) ? value : 0;
        }

        private decimal GetCellValueAsDecimal(ExcelRange cell, bool isPercentage = false)
        {
            string? cellValue = cell?.Value?.ToString();
            if (isPercentage && !string.IsNullOrEmpty(cellValue))
            {
                cellValue = cellValue.Replace("%", "");
            }
            return decimal.TryParse(cellValue, out decimal value) ? value : 0;
        }

        private string? GetCellValueAsString(ExcelRange cell)
        {
            return cell?.Value?.ToString();
        }
    }
}
