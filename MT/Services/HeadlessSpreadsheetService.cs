using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;

namespace MT.Services
{
    public class HeadlessSpreadsheetService
    {
        const string _connectionString = "Server=(localdb)\\mssqllocaldb;Database=MT;Trusted_Connection=True;MultipleActiveResultSets=true";
        Dictionary<string, string> _metadata;

        public Dictionary<string, string[,]> Process()
        {
            _metadata = RetrieveMetadata();

            var data = new List<DataItem>();

            // Iterate through tables in metadata
            foreach (var table in _metadata.Keys)
            {
                // SQL query to retrieve data from each table
                string query = $"SELECT datasource, * FROM {table}";
                if (table == null) continue;

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var sheetName = _metadata[table];
                            var dataSource = JsonConvert.DeserializeObject<Dictionary<string, int[]>>(reader["datasource"].ToString());
                            if (dataSource == null) continue;
                            var row = dataSource.Values.First()[0];
                            var column = dataSource.Values.First()[1];

                            // Check if the sheetKey matches the expected sheet name
                            if (dataSource.Keys.First() != sheetName) return null;

                            var counter = 0;
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (reader.GetName(i) == "datasource" || reader.GetName(i) == "ID") continue;
                                var item = new DataItem
                                {
                                    SheetName = sheetName,
                                    Row = row,
                                    Column = column + counter,
                                    Value = reader[i].ToString() ?? ""
                                };
                                counter++;
                                data.Add(item);
                            }
                        }
                    }
                }
            }
            var result = BuildDictionary(data);
            new DbService().CreateConfigTableJson(result);

            return result;
        }

        static Dictionary<string, string> RetrieveMetadata()
        {
            // SQL query to retrieve metadata from the database
            string metadataQuery = "SELECT TableName, SheetName FROM Metadata";

            // Dictionary to store the retrieved metadata
            var metadata = new Dictionary<string, string>();

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(metadataQuery, connection))
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string tableName = reader["TableName"].ToString();
                        string sheetName = reader["SheetName"].ToString();

                        // Add the entry to the metadata dictionary
                        metadata[tableName] = sheetName;
                    }
                }
            }

            return metadata;
        }

        static Dictionary<string, string[,]> BuildDictionary(List<DataItem> data)
        {
            var result = new Dictionary<string, string[,]>();

            foreach (var item in data)
            {
                if (!result.ContainsKey(item.SheetName))
                {
                    var i = data.Where(x => x.SheetName == item.SheetName).Max(x => x.Row);
                    var j = data.Where(x => x.SheetName == item.SheetName).Max(x => x.Column);
                    var sheet = new string[i, j];
                    ReplaceNullWithEmptyString(sheet);
                    result[item.SheetName] = sheet;
                }

                // Set the values at the specified coordinate and subsequent columns
                result[item.SheetName][item.Row - 1, item.Column - 1] = item.Value;
            }

            var formulas = new DbService().GetFormula();
            if (!formulas.Any()) return result;
            var rowMax = formulas.Max(x => x.i);
            result["Formulas"] = new string[rowMax, 2];

            foreach (var formula in formulas)
            {
                result["Formulas"][formula.i - 1, 1] = formula.context;
                if (string.IsNullOrEmpty(formula.formula)) result["Formulas"][formula.i - 1, 0] = formula.result;
                else result["Formulas"][formula.i - 1, 0] = "=" + formula.formula;
            }

            return result;
        }

        public static void ReplaceNullWithEmptyString(string[,] array)
        {
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    if (array[i, j] == null)
                    {
                        array[i, j] = "";
                    }
                }
            }
        }

        // Class representing a data item with coordinates and values
        class DataItem
        {
            public string SheetName { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public string Value { get; set; }
        }
    }

}
