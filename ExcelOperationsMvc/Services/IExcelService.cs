using ExcelOperationsMvc.Models;

namespace ExcelOperationsMvc.Services;

public interface IExcelService
{
    Task<(MemoryStream, string, string)> ExportToExcel(List<User> users);
    Task<List<User>> ImportFromExcel(IFormFile formFile);
    (MemoryStream,string,string) ExcelExportByDataTable();
}