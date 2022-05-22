using ExcelOperationsMvc.Models;

namespace ExcelOperationsMvc.Services;

public class UserService : IUserService
{
    public Task<List<User>> GetUserList()
    {
        var usersList = new List<User>()
        {
            new()
            {
                FirstName = "Abylai",
                LastName = "Ayathan",
                Age = 29,
                Address = "Chu"
            },
            new()
            {
                FirstName = "Adilet",
                LastName = "Asankan",
                Age = 32,
                Address = "Karakol"
            },
            new()
            {
                FirstName = "Erbol",
                LastName = "Arlekov",
                Age = 21,
                Address = "Santash"
            },
            new()
            {
                FirstName = "Aizhamal",
                LastName = "Nurlan",
                Age = 26,
                Address = "Santash"
            }
        };
        return Task.FromResult(usersList);
    }
}