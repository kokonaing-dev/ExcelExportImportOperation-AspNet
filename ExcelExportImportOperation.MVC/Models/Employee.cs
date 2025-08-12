using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelExportImportOperation.MVC.Models;

[Table("Employees")]
public class Employee
{
    [Key]
    public int Id { get; set; }
    public string EmployeeID { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Email { get; set; }
    public string PhoneNumber { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public string Gender { get; set; }
    public string Department { get; set; }
    public string JobTitle { get; set; }
    public string Manager { get; set; }
    public DateTime? HireDate { get; set; }
    public decimal? Salary { get; set; }
    public decimal? Bonus { get; set; }
    public string AddressLine1 { get; set; }
    public string AddressLine2 { get; set; }
    public string City { get; set; }
    public string State { get; set; }
    public string ZipCode { get; set; }
    public string Country { get; set; }
    public string EmploymentStatus { get; set; }
    public string WorkLocation { get; set; }
    public string OfficePhone { get; set; }
    public string MobilePhone { get; set; }
    public string EmergencyContactName { get; set; }
    public string EmergencyContactPhone { get; set; }
    public string MaritalStatus { get; set; }
    public string EducationLevel { get; set; }
    public string Certifications { get; set; }
    public string Languages { get; set; }
    public string Notes { get; set; }
}
