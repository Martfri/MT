using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Jering.Javascript.NodeJS;
using MT.Services;
using MT.Models;

namespace MT.Controllers
{
    public class FormulaController : Controller
    {
        private readonly DbService _dbService;
        private readonly HeadlessSpreadsheetService _headlessSpreadsheetService;
        private readonly INodeJSService _nodeJSService;

        public FormulaController(DbService dbService, INodeJSService nodeJSService, HeadlessSpreadsheetService headlessSpreadsheetService)
        {
            _nodeJSService = nodeJSService;
            _dbService = dbService;
            _headlessSpreadsheetService = headlessSpreadsheetService;
        }

        [HttpGet]
        public async Task<IActionResult> CalculateFormulas()
        {
            _headlessSpreadsheetService.Process();
            var internalRepresentation = _dbService.GetJsonConfig();

            var deserialized = JsonConvert.DeserializeObject<Dictionary<string, string[][]>>(internalRepresentation);

            _nodeJSService.InvokeFromFileAsync(@".\wwwroot\js\HyperFormulaScript.js", "CalculateFormula", args: new object[] { deserialized });

            Thread.Sleep(1700);
            var formulasfromdb = new DbService().GetFormula();
            return View("Formulas", formulasfromdb);
        }   

        [HttpGet]
        public IActionResult Formulas(IFormCollection form)
        {
            var formulasfromdb = new DbService().GetFormula();
            return View(formulasfromdb);
        }

        // Delete a formula from the preview 
        [HttpGet]
        public IActionResult Delete(string name)
        {
            try
            {
                _dbService.DeleteFormula(name);
            }
            catch (Exception)
            {
            }
            var formulasfromdb = new DbService().GetFormula();
            return View("Formulas", formulasfromdb);
        }
    }
}