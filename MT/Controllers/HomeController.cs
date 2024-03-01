using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Microsoft.AspNetCore.Authorization;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using System.Data;
using MT.Models;
using MT.Services;

namespace MT.Controllers
{

    public class HomeController : Controller
    {
        // Path where Excel doucment is temporarily stored
        private static string filePath = ".\\wwwroot\\Uploads\\Test.xlsx";
        private ExcelService _spreadsheetService;
        private readonly DbService _dbService;


        public HomeController(DbService dbService, ExcelService excelService)
        {
            _dbService = dbService;
            _spreadsheetService = excelService;
        }
      
      
        [Authorize]
        [HttpGet]
        public IActionResult Index(IFormCollection form)
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcelFile(IFormFile FormFile)
        {
            //get file name
            var filename = ContentDispositionHeaderValue.Parse(FormFile.ContentDisposition).FileName.Trim('"');

            //get path
            var MainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");

            if (Directory.Exists(MainPath))
            {
                Directory.Delete(MainPath, true);
            }

            var filePath = Path.Combine(MainPath, FormFile.FileName);

            string extension = Path.GetExtension(filename);

            string conString = string.Empty;

            // Get extension
            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                    break;
            }

            // Check extension
            if (extension != ".xlsx")
            {
                ViewBag.Message = "Uploaded file is not an xlsx document";
            }
            else
            {
                ViewBag.Message = "File uploaded";
            }

            // Create directory "Uploads" if it doesn't exists
            if (!Directory.Exists(MainPath))
            {
                Directory.CreateDirectory(MainPath);
            }

            // Get file path 
            filePath = Path.Combine(MainPath, "Test.xlsx");

            using (System.IO.Stream stream = new FileStream(filePath, FileMode.Create))
            {
                await FormFile.CopyToAsync(stream);
            }

            var formulasheet = new ExcelService().wss.Where(ws => ws.Name == "Formulas").FirstOrDefault();

            if (formulasheet == null)
            {
                ViewBag.Message = "Uploaded document does not include a Formulas sheet";
                return View("Index");
            }

            if (formulasheet.Cells[1, 1].Value.ToString() != "Formula" ||
                formulasheet.Cells[1, 2].Value.ToString() != "Name")
            {
                ViewBag.Message = "Uploaded document does not have a valid format of the formulas sheet";
                return View("Index");
            }
            var tableInfos = _spreadsheetService.tableInfos;
            
            TempData.Put("tables", tableInfos);
            TempData.Keep("tables");
            return View("Index");
        }     
    }


    // Helper class to store and retrieve temporary data
    public static class TempDataExtensions
    {
        public static void Put<T>(this ITempDataDictionary tempData, string key, T value) where T : class
        {
            tempData[key] = JsonConvert.SerializeObject(value);
        }

        public static T Get<T>(this ITempDataDictionary tempData, string key) where T : class
        {
            object o;
            tempData.TryGetValue(key, out o);
            return o == null ? null : JsonConvert.DeserializeObject<T>((string)o);
        }
    }
}