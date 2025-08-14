using ExcelExportImportOperation.MVC.Models;

namespace ExcelExportImportOperation.MVC.Services;

public interface IEmployeeService
{
    List<Employee> GetEmployees();
    Task SaveToDatabaseAsync(List<DataRowModel> data);
}

public class EmployeeService : IEmployeeService
{
    private readonly AppDbContext _context;

    public EmployeeService(AppDbContext context)
    {
        _context = context;
    }

    public List<Employee> GetEmployees()
    {
        return _context.Employees.ToList();
    }

    public async Task SaveToDatabaseAsync(List<DataRowModel> data)
    {
        var employees = data.Select(d => new Employee
        {
            EmployeeID = d.EmployeeID,
            FirstName = d.FirstName,
            LastName = d.LastName,
            Email = d.Email,
            PhoneNumber = d.PhoneNumber,
            DateOfBirth = DateTime.TryParse(d.DateOfBirth, out var dob) ? dob : null,
            Gender = d.Gender,
            Department = d.Department,
            JobTitle = d.JobTitle,
            Manager = d.Manager,
            HireDate = DateTime.TryParse(d.HireDate, out var hire) ? hire : null,
            Salary = decimal.TryParse(d.Salary, out var sal) ? sal : null,
            Bonus = decimal.TryParse(d.Bonus, out var bonus) ? bonus : null,
            AddressLine1 = d.AddressLine1,
            AddressLine2 = d.AddressLine2,
            City = d.City,
            State = d.State,
            ZipCode = d.ZipCode,
            Country = d.Country,
            EmploymentStatus = d.EmploymentStatus,
            WorkLocation = d.WorkLocation,
            OfficePhone = d.OfficePhone,
            MobilePhone = d.MobilePhone,
            EmergencyContactName = d.EmergencyContactName,
            EmergencyContactPhone = d.EmergencyContactPhone,
            MaritalStatus = d.MaritalStatus,
            EducationLevel = d.EducationLevel,
            Certifications = d.Certifications,
            Languages = d.Languages,
            Notes = d.Notes
        }).ToList();

        _context.Employees.AddRange(employees);
        await _context.SaveChangesAsync();
    }
}
