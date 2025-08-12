using ClosedXML.Excel;
using ExcelExportImportOperation.MVC.Models;

namespace ExcelExportImportOperation.MVC.Services;

public interface IExcelImportService
{
    List<DataRowModel> ReadExcel(IFormFile file);
}

public class ExcelImportService : IExcelImportService
{
    public List<DataRowModel> ReadExcel(IFormFile file)
    {
        List<DataRowModel> rows = new();

        using (var stream = file.OpenReadStream())
        using (var workbook = new XLWorkbook(stream))
        {
            var ws = workbook.Worksheet(1);
            var usedRange = ws.RangeUsed();
            if (usedRange == null) throw new Exception("Excel file is empty.");

            var headerRow = usedRange.FirstRow();
            Dictionary<string, int> headerMap = new();
            int colIndex = 1;

            foreach (var cell in headerRow.Cells())
            {
                string header = cell.GetValue<string>().Trim().ToLower();
                if (!headerMap.ContainsKey(header))
                    headerMap[header] = colIndex;
                colIndex++;
            }

            string GetCellValue(IXLRangeRow row, string colName)
            {
                colName = colName.ToLower();
                if (headerMap.ContainsKey(colName))
                {
                    return row.Cell(headerMap[colName]).GetValue<string>().Trim();
                }
                return null;
            }

            foreach (var row in usedRange.RowsUsed().Skip(1))
            {
                rows.Add(new DataRowModel
                {
                    EmployeeID = GetCellValue(row, "employeeid"),
                    FirstName = GetCellValue(row, "firstname"),
                    LastName = GetCellValue(row, "lastname"),
                    Email = GetCellValue(row, "email"),
                    PhoneNumber = GetCellValue(row, "phonenumber"),
                    DateOfBirth = GetCellValue(row, "dateofbirth"),
                    Gender = GetCellValue(row, "gender"),
                    Department = GetCellValue(row, "department"),
                    JobTitle = GetCellValue(row, "jobtitle"),
                    Manager = GetCellValue(row, "manager"),
                    HireDate = GetCellValue(row, "hiredate"),
                    Salary = GetCellValue(row, "salary"),
                    Bonus = GetCellValue(row, "bonus"),
                    AddressLine1 = GetCellValue(row, "addressline1"),
                    AddressLine2 = GetCellValue(row, "addressline2"),
                    City = GetCellValue(row, "city"),
                    State = GetCellValue(row, "state"),
                    ZipCode = GetCellValue(row, "zipcode"),
                    Country = GetCellValue(row, "country"),
                    EmploymentStatus = GetCellValue(row, "employmentstatus"),
                    WorkLocation = GetCellValue(row, "worklocation"),
                    OfficePhone = GetCellValue(row, "officephone"),
                    MobilePhone = GetCellValue(row, "mobilephone"),
                    EmergencyContactName = GetCellValue(row, "emergencycontactname"),
                    EmergencyContactPhone = GetCellValue(row, "emergencycontactphone"),
                    MaritalStatus = GetCellValue(row, "maritalstatus"),
                    EducationLevel = GetCellValue(row, "educationlevel"),
                    Certifications = GetCellValue(row, "certifications"),
                    Languages = GetCellValue(row, "languages"),
                    Notes = GetCellValue(row, "notes")
                });
            }
        }

        return rows;
    }
}