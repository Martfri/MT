using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using Newtonsoft.Json;
using System.Data;
using Formatting = Newtonsoft.Json.Formatting;
using MT.Models;
using MT.Services;

namespace MT.Controllers
{
    public class TableController : Controller
    {
        private readonly DbService _dbService;

        public TableController(DbService dbService)
        {
            _dbService = dbService;
        }      

        // Get table view
        [Authorize]
        [HttpGet]
        public IActionResult TableView()
        {
            List<Tableinfo> tables = _dbService.GetTableInfoFromDatabase();
            return View(tables);
        }     


        // Get preview
        [HttpGet]
        public IActionResult Preview(string name)
        {
            var table = _dbService.RetrieveTable(name);
            return View(table);
        }

        // Get Edit view
        [HttpGet]
        public IActionResult Edit(int name, string tablename)
        {
            var table = _dbService.RetrieveTable(tablename, name);
            Dictionary<string, Type> names = new Dictionary<string, Type>();

            foreach (DataColumn c in table.Columns)
            {
                names.Add(c.ColumnName, c.DataType);
            }

            TempData.Put("names", names);
            return View(table);
        }

        [HttpGet]
        public IActionResult Insert(string tablename)
        {
            var table = _dbService.RetrieveTable(tablename);
            Dictionary<string, Type> names = new Dictionary<string, Type>();

            foreach (DataColumn c in table.Columns)
            {
                names.Add(c.ColumnName, c.DataType);
            }

            TempData.Put("names", names);

            return View(table);
        }

        [HttpGet]
        public IActionResult DeleteRow(int id, string tablename)
        {
            var datasource = _dbService.RetrieveJson(tablename, id);
            _dbService.DeleteRowById(tablename, id);
            var table = _dbService.RetrieveTable(tablename);

            var value = datasource[datasource.Keys.FirstOrDefault()];
            var newdic = new List<Tuple<int, int, string, string>>();
            var counter = 0;

            foreach (DataColumn c in table.Columns)
            {
                Tuple<int, int, string, string> myTuple = new Tuple<int, int, string, string>(value[0] - 1, value[1] - 1 + counter, "", datasource.Keys.FirstOrDefault());
                newdic.Add(myTuple);
                counter++;
            }
            return View("Preview", table);
        }

        [HttpPost]
        public IActionResult InsertRow(IFormCollection form)
        {
            Dictionary<string, Type> names = TempData.Get<Dictionary<string, Type>>("names");
            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            foreach (var name in names)
            {
                if (name.Key == "ID") continue;
                if (name.Key == "datasource") continue;
                if (name.Value == null) continue;

                if (name.Value == typeof(int)) keyValuePairs.Add(name.Key, int.Parse(form[name.Key]));
                else if (name.Value == typeof(double)) keyValuePairs.Add(name.Key, double.Parse(form[name.Key]));
                else if (name.Value == typeof(float)) keyValuePairs.Add(name.Key, float.Parse(form[name.Key]));
                else if (name.Value == typeof(decimal)) keyValuePairs.Add(name.Key, decimal.Parse(form[name.Key]));
                else keyValuePairs.Add(name.Key, form[name.Key]);
            }
            var tableName = form["TableName"];

            var key = _dbService.RetrieveJson(tableName);
            var value = key[key.Keys.FirstOrDefault()];
            key[key.Keys.FirstOrDefault()][0] = value[0] + 1;
            _dbService.AddRowToTable(tableName, keyValuePairs, key);

            List<string> row = new List<string>();
            for (int i = 1; i < value[1]; i++)
            {
                row.Add("");
            }
            foreach (var kv in keyValuePairs)
            {
                row.Add(kv.Value.ToString());
            }

            var table = _dbService.RetrieveTable(tableName);

            return View("Preview", table);
        }

        [HttpGet]
        public IActionResult Update(DataRow name)
        {
            return View();
        }
     
        // Set a new name for a table
        [HttpPost]
        public IActionResult EditTable(IFormCollection form)
        {
            Dictionary<string, Type> names = TempData.Get<Dictionary<string, Type>>("names");
            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            var tableName = form["TableName"];
            int counter = 0;

            foreach (var name in names)
            {
                var a = form[name.Key];

                if (string.IsNullOrEmpty(a)) continue;
                else if (name.Key == "datasource") continue;

                if (name.Value == typeof(int)) keyValuePairs.Add(name.Key, int.Parse(form[name.Key]));
                else if (name.Value == typeof(double)) keyValuePairs.Add(name.Key, double.Parse(form[name.Key]));
                else if (name.Value == typeof(float)) keyValuePairs.Add(name.Key, float.Parse(form[name.Key]));
                else if (name.Value == typeof(decimal)) keyValuePairs.Add(name.Key, decimal.Parse(form[name.Key]));
                else keyValuePairs.Add(name.Key, form[name.Key]);
                counter++;
            }

            if (counter <= 1)
            {
                var tableOld = _dbService.RetrieveTable(tableName);

                return View("Preview", tableOld);
            }

            _dbService.UpdateTable(tableName, keyValuePairs);

            var datasorce = _dbService.RetrieveJson(tableName, int.Parse(form["ID"]));

            var value = datasorce[datasorce.Keys.FirstOrDefault()];
         
            List<string> row = new List<string>();
            for (int i = 1; i < value[1]; i++)
            {
                row.Add("");
            }
            foreach (var kv in keyValuePairs)

            {
                if (kv.Key == "ID") continue;
                row.Add(kv.Value.ToString());
            }

            var table = _dbService.RetrieveTable(tableName);

            return View("Preview", table);
        }

        // Delete a table from the preview 
        [HttpGet]
        public IActionResult Delete(string name)
        {
            try
            {
                _dbService.DeleteRowFromMetadata(name);
                _dbService.DeleteTable(name);
            }
            catch (Exception)
            {
            }

            List<Tableinfo> tables = _dbService.GetTableInfoFromDatabase();
            return View("TableView",tables);
        }

        // Export data model to a JSON file
        [HttpGet]
        public IActionResult Export()
        {
            var jsonData = JsonConvert.SerializeObject(TempData.Get<List<Table>>("tables"), Formatting.Indented);

            string fileName = "jsonExport.json";
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(jsonData);

            var content = new MemoryStream(bytes);
            return File(content, "application/json", fileName);
        }

        [HttpPost]
        public void EditTableTest([FromBody] Foo data)
        {
            var id = new Random().Next(2, 250);
            var keyValuePairs = new Dictionary<string, object>
    {
        { "Spieler", data.Spieler },
        { "Position", data.Position },
        { "Monatsgehalt", data.Monatsgehalt },
        { "ID", id }
    };
            var tableName = data.tablename;
            _dbService.UpdateTable(tableName, keyValuePairs);
        }

        [HttpPost]
        public void InsertTableTest([FromBody] Foo data)
        {
            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            keyValuePairs.Add("Spieler", data.Spieler);
            keyValuePairs.Add("Position", data.Position);
            keyValuePairs.Add("Monatsgehalt", data.Monatsgehalt);

            var tableName = data.tablename;
            var json = _dbService.RetrieveJson(tableName);
            var value = json[json.Keys.FirstOrDefault()];
            json[json.Keys.FirstOrDefault()][0] = value[0] + 1;
            _dbService.AddRowToTable(tableName, keyValuePairs, json);
        }

        [HttpPost]
        public void DeleteTableTest([FromBody] Foo data)
        {
            var tablename = data.tablename;
            var id = new Random().Next(5, 9980);
            var db = _dbService;
            db.DeleteRowById(tablename, id);
        }

        [HttpPost]
        public void ReadTableTest()
        {
            var db = _dbService;
            var formulasfromdb = db.GetFormula();
        }

        public class Foo
        {
            public object Spieler { get; set; }
            public string tablename { get; set; }
            public object Position { get; set; }
            public object Monatsgehalt { get; set; }

        }
    }
}