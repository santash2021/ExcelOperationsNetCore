using System.Data;
using System.Drawing;
using System.Reflection.Metadata;
using ExcelOperationsMvc.Commons;
using ExcelOperationsMvc.Models;
using OfficeOpenXml;

namespace ExcelOperationsMvc.Services;

public class ExcelService: IExcelService
{
    private readonly IUserService _userService;

    public ExcelService(IUserService userService)
    {
        _userService = userService;
    }

    public Task<(MemoryStream, string, string)> ExportToExcel(List<User> users)
    {
        var memoryStream = new MemoryStream();
        var xlPackage = new ExcelPackage(memoryStream);
        var worksheet = xlPackage.Workbook.Worksheets.Add("UsersList");

        worksheet.Cells["A1"].Value = "Sample user's list";

        using var worksheetCells = worksheet.Cells["A1:D1"];
        worksheetCells.Merge = true;
        worksheetCells.Style.Font.Color.SetColor(Color.Chartreuse);
        worksheetCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        worksheetCells.Style.Fill.BackgroundColor.SetColor(Color.Fuchsia);

        worksheet.Cells["A4"].Value = "FirstName";
        worksheet.Cells["B4"].Value = "LastName";
        worksheet.Cells["C4"].Value = "Age";
        worksheet.Cells["D4"].Value = "Address";

        worksheet.Cells["A4:D4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        worksheet.Cells["A4:D4"].Style.Fill.BackgroundColor.SetColor(Color.Aqua);

        var startRow = 5;
        var row = startRow;

        foreach (var user in users)
        {
            worksheet.Cells[row, 1].Value = user.FirstName;
            worksheet.Cells[row, 2].Value = user.LastName;
            worksheet.Cells[row, 3].Value = user.Age;
            worksheet.Cells[row, 4].Value = user.Address;

            row++;
        }

        xlPackage.Workbook.Properties.Title = "Demo excel";
        xlPackage.Workbook.Properties.Author = "Ermek";

        xlPackage.Save();

        memoryStream.Position = 0;
        return Task.FromResult((memoryStream,Constants.Excel,$"{nameof(User)}.{Constants.ExcelFileExtensionType}"));
    }

    public Task<List<User>> ImportFromExcel(IFormFile formFile)
    {
        var stream = formFile.OpenReadStream();
        var userList = new List<User>();
        try
        {
            using var package = new ExcelPackage(stream);
            var workSheet = package.Workbook.Worksheets.FirstOrDefault();
            var rowCount = workSheet?.Dimension.Rows;

            for (var row = 5; row <= rowCount; row++)
            {
                var firstName = workSheet?.Cells[row, 1]?.Value?.ToString();
                var lastName = workSheet?.Cells[row, 2]?.Value?.ToString();
                var age = workSheet?.Cells[row, 3]?.Value;
                var address = workSheet?.Cells[row, 4]?.Value?.ToString();

                var user = new User
                {
                    FirstName = firstName,
                    LastName = lastName,
                    Age = Convert.ToInt32(age ?? 0),
                    Address = address
                };
                userList.Add(user);
            }

            return Task.FromResult(userList);
        }
        catch (Exception e)
        {
            return Task.FromResult(new List<User>());
        }
    }


    public (MemoryStream,string,string) ExcelExportByDataTable()
    {
        var ms = new MemoryStream();
        
        var xlPackage = new ExcelPackage(ms);
        var ds = new DataSet();
        ds.Tables.Add(GetTable("User").Result);
        var worksheet = xlPackage.Workbook.Worksheets.Add("UsersList");
        
        worksheet.Cells["A1"].LoadFromDataTable(ds.Tables[0],true);
        xlPackage.SaveAs(ms);
        ms.Position = 0;
        return (ms,Constants.Excel,$"{nameof(User)}.{Constants.ExcelFileExtensionType}");
    }
    
    private async Task<DataTable> GetTable(string tableName)
    {
        List<User> users = await _userService.GetUserList();

        DataTable table = new DataTable { TableName = tableName };

        table.Columns.Add("FirstName", typeof(string));
        table.Columns.Add("LastName", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Columns.Add("Address", typeof(string));

        users.ForEach(x =>
        {
            table.Rows.Add(x.FirstName, x.LastName, x.Age, x.Address);
        });

        return table;
    }
}