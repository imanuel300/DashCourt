using OfficeOpenXml;
using DashCourtApi.Models;
using System.Text.Json; // Add for JSON serialization

namespace DashCourtApi.Services
{
    public class ExcelDataService
    {
        private readonly string _excelFilesPath;
        private readonly bool _useMockData;
        private readonly string _jsonOutputPath; // New field for JSON output path

        public ExcelDataService(string excelFilesPath, bool useMockData, string jsonOutputPath)
        {
            _excelFilesPath = excelFilesPath;
            _useMockData = useMockData;
            _jsonOutputPath = jsonOutputPath; // Initialize new field
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
                        Inventory = GetCellValueAsInt(worksheet.Cells[row, 12]),
                        Year = GetCellValueAsInt(worksheet.Cells[row, 10]),
                        CR = GetCellValueAsDecimal(worksheet.Cells[row, 9], true), // true for percentage
                        Opened = GetCellValueAsInt(worksheet.Cells[row, 7]),
                        Closed = GetCellValueAsInt(worksheet.Cells[row, 8]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 5]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 4]),
                        District = GetCellValueAsString(worksheet.Cells[row, 3]),
                        Cycle = GetCellValueAsString(worksheet.Cells[row, 11]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 1]) // Corrected mapping to column 1
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
                        Year = GetCellValueAsInt(worksheet.Cells[row, 10]),
                        DateCycle = null, // Not directly available in the provided Excel data
                        AverageDays = GetCellValueAsDecimal(worksheet.Cells[row, 8]),
                        NumberOfCases = 0, // Not directly available in the provided Excel data
                        TotalDays = GetCellValueAsInt(worksheet.Cells[row, 9]), // Map 'סהכ_ימים'
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 5]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 4]),
                        District = GetCellValueAsString(worksheet.Cells[row, 3]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 1]) // ערכאה
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
                        Year = GetCellValueAsInt(worksheet.Cells[row, 10]),
                        DelaysPercentage = GetCellValueAsDecimal(worksheet.Cells[row, 9], true),
                        Rejected = GetCellValueAsInt(worksheet.Cells[row, 8]),
                        Hearings = GetCellValueAsInt(worksheet.Cells[row, 7]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 5]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 4]),
                        District = GetCellValueAsString(worksheet.Cells[row, 3]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 1])
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
                    new Inv3Model { DaysInSystem = 300, Average = 250.5m, CurrentInventory = 1000, Year = 2023, YearAndMonth = "2023-01", CaseType = "MockInv3", Procedure = "MockProc", Court = "MockCourt", District = "MockDist", OriginalCycle = "MockOrigCycle" },
                    new Inv3Model { DaysInSystem = 320, Average = 260.0m, CurrentInventory = 1100, Year = 2023, YearAndMonth = "2023-02", CaseType = "MockInv32", Procedure = "MockProc2", Court = "MockCourt2", District = "MockDist2", OriginalCycle = "MockOrigCycle2" }
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
                        DaysInSystem = GetCellValueAsInt(worksheet.Cells[row, 11]),
                        Average = GetCellValueAsDecimal(worksheet.Cells[row, 10]),
                        CurrentInventory = GetCellValueAsInt(worksheet.Cells[row, 9]),
                        Year = GetCellValueAsInt(worksheet.Cells[row, 8]),
                        YearAndMonth = GetCellValueAsString(worksheet.Cells[row, 7]),
                        CaseType = GetCellValueAsString(worksheet.Cells[row, 6]),
                        Procedure = GetCellValueAsString(worksheet.Cells[row, 5]),
                        Court = GetCellValueAsString(worksheet.Cells[row, 4]),
                        District = GetCellValueAsString(worksheet.Cells[row, 3]),
                        OriginalCycle = GetCellValueAsString(worksheet.Cells[row, 1]) // Corrected mapping to column 1
                    });
                }
            }
            return data;
        }

        public void GenerateJsonFiles()
        {
            if (_useMockData) return; // Don't generate JSON if using mock data

            // Ensure the directory exists
            if (!Directory.Exists(_jsonOutputPath))
            {
                Directory.CreateDirectory(_jsonOutputPath);
            }

            // Generate CR data JSON
            var crData = GetCRData();
            string crJson = JsonSerializer.Serialize(crData, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(_jsonOutputPath, "CR.json"), crJson);

            // Generate AVGO data JSON
            var avgoData = GetAVGOData();
            string avgoJson = JsonSerializer.Serialize(avgoData, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(_jsonOutputPath, "AVGO.json"), avgoJson);

            // Generate SIT data JSON
            var sitData = GetSITData();
            string sitJson = JsonSerializer.Serialize(sitData, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(_jsonOutputPath, "SIT.json"), sitJson);

            // Generate Inv3 data JSON
            var inv3Data = GetInv3Data();
            string inv3Json = JsonSerializer.Serialize(inv3Data, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(_jsonOutputPath, "Inv3.json"), inv3Json);
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
