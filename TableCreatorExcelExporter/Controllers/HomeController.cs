using Ganss.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using TableCreatorExcelExporter.Models;

namespace TableCreatorExcelExporter.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View(new IndexViewModel());
        }

        [HttpPost]
        public IActionResult Index(IndexViewModel model)
        {
            string[] lines = model.TextData.Split("\r\n");
            string[] columns = lines[0].Split(";");

            model.Columns.AddRange(columns);

            for (int i = 1; i < lines.Length; i++)
            {
                List<string> oneRowData = new List<string>();
                oneRowData.AddRange(lines[i].Split(";"));

                model.Rows.Add(oneRowData);
            }

            return View(model);
        }
        
        [HttpPost]
        public IActionResult ExcelExport(string text)
        {
            string[] lines = text.Split("\r\n");
            string[] columns = lines[0].Split(";");

            DataTable dt = new DataTable("datatable");
            columns.ToList().ForEach(c => dt.Columns.Add(c));

            //foreach (string c in columns)
            //{
            //    dt.Columns.Add(c);
            //}

            for (int i = 1; i < lines.Length; i++)
            {
                dt.Rows.Add(lines[i].Split(";"));
            }

            ExcelMapper mapper = new ExcelMapper();
            mapper.Save("wwwroot/data.xlsx", dt, "data");

            return File("~/data.xlsx", "application/vnd.ms-excel","excel_export.xlsx");
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
}