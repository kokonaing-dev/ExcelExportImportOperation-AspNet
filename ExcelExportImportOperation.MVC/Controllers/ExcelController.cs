using ClosedXML.Excel;
using ExcelExportImportOperation.MVC.Models;
using ExcelExportImportOperation.MVC.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelExportImportOperation.MVC.Controllers;

public class ExcelController : Controller
{

    private readonly IExcelImportService _excelImportService;
    private readonly IExcelExportService _excelExportService;
    private readonly IEmployeeService _employeeService;

    public ExcelController(IExcelImportService excelImportService, IEmployeeService employeeService, IExcelExportService excelExportService)
    {
        _excelImportService = excelImportService;
        _employeeService = employeeService;
        _excelExportService = excelExportService;
    }

    [HttpGet]
    public ActionResult Index()
    {
        return View();
    }


    [HttpPost]
    public async Task<IActionResult> Import(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            ViewBag.Error = "Please upload an Excel file.";
            return View("Index");
        }

        try
        {
            var rows = _excelImportService.ReadExcel(file);
            await _employeeService.SaveToDatabaseAsync(rows);

            ViewBag.Message = $"Successfully imported {rows.Count} rows.";
        }
        catch (Exception ex)
        {
            ViewBag.Error = ex.Message;
        }

        return View("Index");
    }

    [HttpGet]
    public IActionResult Export()
    {
        var employees = _employeeService.GetEmployees();
        var fileContent = _excelExportService.ExportEmployees(employees);

        return File(fileContent,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Employees.xlsx");
    }
}
