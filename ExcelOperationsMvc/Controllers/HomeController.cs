using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ExcelOperationsMvc.Models;
using ExcelOperationsMvc.Services;
using OfficeOpenXml;

namespace ExcelOperationsMvc.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly IUserService _userService;
    private readonly IExcelService _excelService;

    public HomeController(ILogger<HomeController> logger, IUserService userService, IExcelService excelService)
    {
        _logger = logger;
        _userService = userService;
        _excelService = excelService;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpGet]
    public async Task<IActionResult> ExportToExcel()
    {
        var userList = await _userService.GetUserList();

        var (excelFileMemoryStream,contentType,fileName) = await _excelService.ExportToExcel(userList);

        return File(excelFileMemoryStream, contentType, fileName);
    }
    
    [HttpGet]
    public IActionResult ExportToExcelByDataTable()
    {
        var (excelFileMemoryStream,contentType,fileName) = _excelService.ExcelExportByDataTable();

        return File(excelFileMemoryStream, contentType, fileName);
    }

    [HttpGet]
    public IActionResult ReadFromExcel()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> ReadFromExcel(IFormFile formFile)
    {
        if (!ModelState.IsValid) return BadRequest();
        if (formFile.Length <= 0) return BadRequest();

        var userList = await _excelService.ImportFromExcel(formFile);
        return View("Users", userList);
    }

    public IActionResult Users(List<User> users)
    {
        return View(users);
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}