using ASPExcel.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace ASPExcel.Controllers
{
    public class UsersController :Controller
    {
        public IActionResult Index()
        {
            var users = GetUsersList();

            return View(users);
        }

        public IActionResult ExportToExcel()
        {
            var users = GetUsersList();

            var stream = new MemoryStream();

            using (var xlPackage=new ExcelPackage(stream))
            {
                var worksheet = xlPackage.Workbook.Worksheets.Add("Users");

                var customStyle = xlPackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                customStyle.Style.Font.UnderLine = true;
                //customStyle.Style.Font.Color.SetColor(Color.Red);

                var starRow = 5;
                var row = starRow;

                worksheet.Cells["A1"].Value = "Test Export";

                using(var r=worksheet.Cells["A1:C1"])
                {
                    r.Merge = true;
                    //r.Style.Fill.BackgroundColor.SetColor(  );
                    
                }
                worksheet.Cells["A4"].Value = "Name";
                worksheet.Cells["B4"].Value = "Email";
                worksheet.Cells["C4"].Value = "Phone";

                row = 5;
                foreach (var user in users)
                {
                    worksheet.Cells[row, 1].Value = user.Name;
                    worksheet.Cells[row, 2].Value = user.Email;
                    worksheet.Cells[row, 3].Value = user.Phone;

                    row++;
                }

                xlPackage.Workbook.Properties.Title = "User lost";
                xlPackage.Workbook.Properties.Author = "Mohamad";

                xlPackage.Save();
            }

            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users.xlsx");

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
