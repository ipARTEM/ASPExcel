using ASPExcel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace ASPExcel.Controllers
{
    public class UsersController :Controller
    {
        public IActionResult Index()
        {
            var users = GetUsersList();

            return View(users);
        }

        private List<User> GetUsersList()
        {
            var users = new List<User>()
            {
                    new User{
                Email="a.khimin@yandex.ru",
                    Name="Artem",
                    Phone="8888888",
            },
                                        new User{
                Email="a.khimin@yandex.ru",
                    Name="Andrey",
                    Phone="9999999",
            },
                                                            new User{
                Email="a.khimin@yandex.ru",
                    Name="Petya",
                    Phone="7777777",
            },
            };

            return users;


        }
        

    }
}
