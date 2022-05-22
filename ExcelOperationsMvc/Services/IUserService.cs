using ExcelOperationsMvc.Models;

namespace ExcelOperationsMvc.Services;

public interface IUserService
{
    Task<List<User>> GetUserList();
}