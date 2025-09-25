using OfficeOpenXml;
using DashCourtApi.Models;

namespace DashCourtApi.Services
{
    public class ExcelDataService
    {
        private readonly string _excelFilesPath;

        public ExcelDataService(string excelFilesPath)
        {
            _excelFilesPath = excelFilesPath;
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public List<CRModel> GetCRData(string fileName = "CR.xlsx")
        {
            var data = new List<CRModel>();
            var filePath = Path.Combine(_excelFilesPath, fileName);

            if (!File.Exists(filePath))
            {
                // Handle file not found error, e.g., throw exception or return empty list
                return data;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Assuming header in row 1
                {
                    data.Add(new CRModel
                    {
                        Inventory = int.TryParse(worksheet.Cells[row, 2].Value?.ToString(), out int inventory) ? inventory : 0,
                        Year = int.TryParse(worksheet.Cells[row, 3].Value?.ToString(), out int year) ? year : 0,
                        CR = decimal.TryParse(worksheet.Cells[row, 4].Value?.ToString().Replace("%", ""), out decimal cr) ? cr : 0,
                        Opened = int.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out int opened) ? opened : 0,
                        Closed = int.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out int closed) ? closed : 0,
                        CaseType = worksheet.Cells[row, 7].Value?.ToString(),
                        Procedure = worksheet.Cells[row, 8].Value?.ToString(),
                        Court = worksheet.Cells[row, 9].Value?.ToString(),
                        District = worksheet.Cells[row, 10].Value?.ToString(),
                        Cycle = worksheet.Cells[row, 11].Value?.ToString(),
                        OriginalCycle = worksheet.Cells[row, 12].Value?.ToString()
                    });
                }
            }
            return data;
        }

        public List<AVGOModel> GetAVGOData(string fileName = "AVGO.xlsx")
        {
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
                        Year = int.TryParse(worksheet.Cells[row, 2].Value?.ToString(), out int year) ? year : 0,
                        DateCycle = worksheet.Cells[row, 3].Value?.ToString(),
                        AverageDays = decimal.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out decimal avgDays) ? avgDays : 0,
                        NumberOfCases = int.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out int numCases) ? numCases : 0,
                        CaseType = worksheet.Cells[row, 6].Value?.ToString(),
                        Procedure = worksheet.Cells[row, 7].Value?.ToString(),
                        Court = worksheet.Cells[row, 8].Value?.ToString(),
                        District = worksheet.Cells[row, 9].Value?.ToString(),
                        OriginalCycle = worksheet.Cells[row, 10].Value?.ToString()
                    });
                }
            }
            return data;
        }

        public List<SITModel> GetSITData(string fileName = "SIT.xlsx")
        {
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
                        Year = int.TryParse(worksheet.Cells[row, 2].Value?.ToString(), out int year) ? year : 0,
                        DelaysPercentage = decimal.TryParse(worksheet.Cells[row, 3].Value?.ToString().Replace("%", ""), out decimal delaysPerc) ? delaysPerc : 0,
                        Rejected = int.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out int rejected) ? rejected : 0,
                        Hearings = int.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out int hearings) ? hearings : 0,
                        CaseType = worksheet.Cells[row, 6].Value?.ToString(),
                        Procedure = worksheet.Cells[row, 7].Value?.ToString(),
                        Court = worksheet.Cells[row, 8].Value?.ToString(),
                        District = worksheet.Cells[row, 9].Value?.ToString(),
                        OriginalCycle = worksheet.Cells[row, 10].Value?.ToString()
                    });
                }
            }
            return data;
        }

        public List<Inv3Model> GetInv3Data(string fileName = "Inv3.xlsx")
        {
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
                        BaseYear = int.TryParse(worksheet.Cells[row, 2].Value?.ToString(), out int baseYear) ? baseYear : 0,
                        DaysInSystem = int.TryParse(worksheet.Cells[row, 3].Value?.ToString(), out int daysInSystem) ? daysInSystem : 0,
                        Average = decimal.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out decimal average) ? average : 0,
                        CurrentInventory = int.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out int currentInventory) ? currentInventory : 0,
                        Year = int.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out int year) ? year : 0,
                        YearAndMonth = worksheet.Cells[row, 7].Value?.ToString(),
                        CaseType = worksheet.Cells[row, 8].Value?.ToString(),
                        Procedure = worksheet.Cells[row, 9].Value?.ToString(),
                        Court = worksheet.Cells[row, 10].Value?.ToString(),
                        District = worksheet.Cells[row, 11].Value?.ToString(),
                        OriginalCycle = worksheet.Cells[row, 12].Value?.ToString()
                    });
                }
            }
            return data;
        }
    }
}
