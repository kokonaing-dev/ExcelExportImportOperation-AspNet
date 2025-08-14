using ClosedXML.Excel;
using ExcelExportImportOperation.MVC.Models;

namespace ExcelExportImportOperation.MVC.Services;

public interface IExcelExportService
{
    byte[] ExportEmployees(IEnumerable<Employee> employees);
}

public class ExcelExportService : IExcelExportService
{
    public byte[] ExportEmployees(IEnumerable<Employee> employees)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Employees");

            // Header
            var headers = typeof(Employee).GetProperties().Select(p => p.Name).ToArray();
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
            }

            // Data
            int row = 2;
            foreach (var emp in employees)
            {
                var props = emp.GetType().GetProperties();
                for (int col = 0; col < props.Length; col++)
                {
                    var rawValue = props[col].GetValue(emp);
                    worksheet.Cell(row, col + 1).Value = rawValue?.ToString() ?? string.Empty;
                }
                row++;
            }

            // Auto-fit
            worksheet.Columns().AdjustToContents();

            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                return stream.ToArray();
            }
        }
    }
}

