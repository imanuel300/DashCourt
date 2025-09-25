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
            // For EPPlus 8 and later, use ExcelPackage.License
            if (ExcelPackage.LicenseContext == LicenseContext.Unspecified) // Check if already set
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial
            }
        }

        public List<CRModel> GetCRData(string fileName = "CR.xlsx")
        {
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
