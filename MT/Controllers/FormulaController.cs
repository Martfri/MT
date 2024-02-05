using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Jering.Javascript.NodeJS;
using MT.Services;

namespace MT.Controllers
{
    public class FormulaController : Controller
    {
        private readonly DbService _dbService;
        private readonly INodeJSService _nodeJSService;

        public FormulaController(DbService dbService, INodeJSService nodeJSService)
        {
            _nodeJSService = nodeJSService;
            _dbService = dbService;
        }

        [HttpGet]
        public IActionResult CalculateFormulas()
        {
            var headlessSpreadsheetService = new HeadlessSpreadsheetService();

            headlessSpreadsheetService.Process();
            var internalRepresentation = _dbService.GetJsonConfig();

            var deserialized = JsonConvert.DeserializeObject<Dictionary<string, string[][]>>(internalRepresentation);

            var result = _nodeJSService.InvokeFromFileAsync(@".\wwwroot\js\HyperFormulaScript.js", "CalculateFormula", args: new object[] { deserialized });

            var formulasfromdb = new DbService().GetFormula();
            return View("Formulas", formulasfromdb);
        }   

        [HttpGet]
        public IActionResult Formulas(IFormCollection form)
        {
            var formulasfromdb = new DbService().GetFormula();
            return View(formulasfromdb);
        }        
    }
}