using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // creating file location for excel file
            //var file = new FileInfo(@"D:\Coding\Visual Studio\Projects\ExcelDemoApp\ExcelDemo.xlsx");
            var file = new FileInfo(@"..\..\..\ExcelDemo.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);

            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.ID} {p.FirstName} {p.LastName}");
            }
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            // return the first worksheet in the excel file
            var worksheet = package.Workbook.Worksheets[0];

            // row 1 is report title, row 2 is header
            int row = 3;
            int col = 1;

            // checking to find last row with data.
            // Since the ID column will always have a number for valid rows, keep processing until no more ID rows
            while (string.IsNullOrWhiteSpace(worksheet.Cells[row,col].Value?.ToString()) == false)
            {
                PersonModel p = new();

                p.ID = int.Parse(worksheet.Cells[row, col].Value.ToString());
                p.FirstName = worksheet.Cells[row, col + 1].Value.ToString();
                p.LastName = worksheet.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;
            }
            return output;
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            // by using the using statement, it will automatically close the package object
            using var package = new ExcelPackage(file);

            // adds a worksheet to the ExcelDemo.xlsx spreadsheet
            var worksheet = package.Workbook.Worksheets.Add("MainReport");

            // stating where the data (with headers) should start. Headers are property names
            var range = worksheet.Cells["A2"].LoadFromCollection(people, true);

            range.AutoFitColumns();

            // Formats for report title
            worksheet.Cells["A1"].Value = "Our Cool Report";
            worksheet.Cells["A1:C1"].Merge = true;
            worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(1).Style.Font.Size = 24;
            worksheet.Row(1).Style.Font.Color.SetColor(Color.Blue);

            // Formats for header 
            worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(2).Style.Font.Bold = true;

            // Formats for data
            worksheet.Column(3).Width = 20;

            await package.SaveAsync();
        }

        /// <summary>
        /// Deletes the Excel file if it currently exists in the directory
        /// </summary>
        /// <param name="file"></param>
        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        /// <summary>
        /// Seed data setup.  Could also set up to read from database or user input from form
        /// </summary>
        /// <returns></returns>
        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { ID = 1, FirstName = "Izabella ", LastName = "Butler" },
                new() { ID = 2, FirstName = "Caleb", LastName = "Horton" },
                new() { ID = 3, FirstName = "Alessandra", LastName = "Perez" },
                new() { ID = 4, FirstName = "Rylie", LastName = "Schultz" },
                new() { ID = 5, FirstName = "Leonel", LastName = "Perez" },
                new() { ID = 6, FirstName = "Adelyn", LastName = "Phillips " },
                new() { ID = 7, FirstName = "Leonardo", LastName = "Perkins" },
                new() { ID = 8, FirstName = "Leonel ", LastName = "Castillo" },
                new() { ID = 9, FirstName = "Isabella", LastName = "Gill" },
                new() { ID = 10, FirstName = "Alyssa", LastName = "Edwards " },
            };

            return output;
        }
    }
}
