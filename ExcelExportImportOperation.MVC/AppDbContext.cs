using ExcelExportImportOperation.MVC.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelExportImportOperation.MVC;

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions options) : base(options)
    {
    }

    public DbSet<Employee> Employees { get; set; }

}
