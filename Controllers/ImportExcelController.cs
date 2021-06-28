using ASPExcel.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;


namespace ASPExcel.Controllers
{
    public class ImportExcelController : Controller
    {
        public IActionResult Index()
        {
            

            return View();
        }

        [HttpGet]
        public IActionResult BatchUserUpload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult BatchUserUpload(IFormFile batchUsers)
        {
            if (ModelState.IsValid)
            {
                if (batchUsers?.Length > 0)
                {
                    // convert to a stream
                    var stream = batchUsers.OpenReadStream();

                    List<User> users = new List<User>();

                    try
                    {
                        using (var package = new ExcelPackage(stream))
                        {
                            var worksheet = package.Workbook.Worksheets.First();
                            var rowCount = worksheet.Dimension.Rows;

                            for (var row = 2; row <= rowCount; row++)
                            {
                                try
                                {
                                    var name = worksheet.Cells[row, 1].Value?.ToString();
                                    var email = worksheet.Cells[row, 2].Value?.ToString();
                                    var phone = worksheet.Cells[row, 3].Value?.ToString();

                                    var user = new User()
                                    {
                                        Email = email,
                                        Name = name,
                                        Phone = phone
                                    };

                                    users.Add(user);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }

                        return View("Index", users);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }

            return View();
        }

        //private List<ImportExcel> GetUsersList()
        //{
        //    var users = new List<ImportExcel>()
        //    {
        //        new ImportExcel{
        //        Email="a.khimin@yandex.ru",
        //            Name="Artem",
        //            Phone="8888888",
        //    },
        //        new ImportExcel{
        //        Email="a.khimin@yandex.ru",
        //            Name="Andrey",
        //            Phone="9999999",
        //    },
        //        new ImportExcel{
        //        Email="a.khimin@yandex.ru",
        //            Name="Petya",
        //            Phone="7777777",
        //    },
        //    };

        //    return users;


        //}

    }
}
