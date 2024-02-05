using System.Data;
using System.Data.SqlClient;
using System.Text;
using Table = MT.Models.Table;
using System.Xml.Linq;
using Newtonsoft.Json;
using MT.Models;


namespace MT.Services
{
    public class DbService
    {
        private Table table;
        const string _connectionString = "Server=(localdb)\\mssqllocaldb;Database=MT;Trusted_Connection=True;MultipleActiveResultSets=true";
        private string tableName;


        public DbService(Table table)
        {
            this.table = table;
            tableName = table.tableName.Replace(' ', '_');
        }

        public DbService()
        {

        }

        public void CreateFormulaTable(List<Formula> formulas)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Formulas' and xtype='U')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"CREATE TABLE Formulas (");
            sb.AppendLine($"[Formula] VARCHAR(1000), [Name] VARCHAR(1000), [Result] VARCHAR(100), [Row] INT");
            sb.AppendLine(")");
            sb.AppendLine("END");
            ExecCommand(sb.ToString());

            DataTable tbl = new DataTable();

            tbl.Columns.Add(new DataColumn("Formula"));
            tbl.Columns.Add(new DataColumn("Name"));
            tbl.Columns.Add(new DataColumn("Result"));
            tbl.Columns.Add(new DataColumn("Row"));


            foreach (Formula formula in formulas)
            {
                DataRow dr = tbl.NewRow();

                //dr["Id"] = 1;
                dr["Formula"] = formula.formula;
                dr["Name"] = formula.name;
                dr["Result"] = formula.result;
                dr["Row"] = formula.i;
                tbl.Rows.Add(dr);
            }

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    objbulk.BulkCopyTimeout = 0;

                    objbulk.DestinationTableName = "Formulas";
                    objbulk.ColumnMappings.Add("Formula", "Formula");
                    objbulk.ColumnMappings.Add("Name", "Name");
                    objbulk.ColumnMappings.Add("Result", "Result");
                    objbulk.ColumnMappings.Add("Row", "Row");

                    try
                    {
                        objbulk.WriteToServer(tbl);
                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        public void AddRowToTable(string tableName, Dictionary<string, object> columnValues, Dictionary<string, int[]> datasource)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    string columnNames = string.Join(", ", columnValues.Keys);
                    string paramNames = string.Join(", ", columnValues.Keys.Select(k => "@" + k));
                    string dataSourceColumnName = "datasource"; // Replace with your actual column name

                    string query = $"INSERT INTO {tableName} ({columnNames}, {dataSourceColumnName}) VALUES ({paramNames}, @datasource)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        foreach (var keyValue in columnValues)
                        {
                            string parameterName = "@" + keyValue.Key;
                            string parameterValue = keyValue.Value.ToString();
                            command.Parameters.AddWithValue(parameterName, parameterValue);
                        }
                        string dataSourceJson = JsonConvert.SerializeObject(datasource);
                        command.Parameters.AddWithValue("@datasource", dataSourceJson);

                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                Console.WriteLine(ex.Message);
            }
        }

        

        //public void UpdateJsonBatch(string key, string[] newRow, int row)
        //{
        //    string connectionString = GetConnectionString("Test"); // Replace with your connection string

        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();

        //        // Start a transaction for batch updates
        //        using (SqlTransaction transaction = connection.BeginTransaction())
        //        {
        //            try
        //            {


        //                // Construct your SQL update statement with parameters
        //                string updateSql = "UPDATE ExcelConfig SET Config = JSON_MODIFY(Config, @JsonPath, JSON_QUERY(@NewJsonValue)) WHERE Id = (SELECT MAX(Id) FROM ExcelConfig);";
        //                // Clean the data by removing backslashes from string elements



        //                string newRowJson = JsonConvert.SerializeObject(newRow, Formatting.None);

        //                using (SqlCommand command = new SqlCommand(updateSql, connection, transaction))
        //                {
        //                    // Set parameters for the update statement
        //                    command.Parameters.AddWithValue("@JsonPath", $@"$.""{key}""[{row}]");
        //                    command.Parameters.AddWithValue("@NewJsonValue", newRowJson);

        //                    // Execute the update statement
        //                    int rowsAffected = command.ExecuteNonQuery();
        //                    Console.WriteLine($"Rows affected: {rowsAffected}");
        //                }


        //                // Commit the transaction after all updates succeed
        //                transaction.Commit();
        //            }
        //            catch (Exception ex)
        //            {
        //                // Handle any exceptions and roll back the transaction if needed
        //                Console.WriteLine($"Error: {ex.Message}");
        //                transaction.Rollback();
        //            }

        //        }
        //    }
        //}

        //private static SemaphoreSlim semaphore = new SemaphoreSlim(1, 1); // Allow only one concurrent task
          

        public List<Formula> GetFormula()
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                string sql = "SELECT Name, Formula, Result, Row FROM Formulas";

                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    //command.Parameters.AddWithValue("Id", 1);
                    var formulas = new List<Formula> { };

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Formula formula = new Formula();
                            formula.name = reader["Name"].ToString();
                            formula.formula = reader["Formula"].ToString();
                            formula.result = reader["Result"].ToString();
                            formula.i = reader["Row"] != DBNull.Value ? (int)reader["Row"] : 0;
                            formulas.Add(formula);
                        }
                    }
                    return formulas;
                }
            }
        }

        public DataTable RetrieveTable(string tableName, int id = 0)
        {
            string tableNamePrepared = tableName.Replace(" ", "_");

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                // Create a SQL command to fetch data from the specified table
                string query = "";
                if (id != 0) { query = $"SELECT * FROM {tableNamePrepared} WHERE ID = {id}"; }
                else { query = $"SELECT * FROM {tableNamePrepared}"; }
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Create a data adapter to fill a DataTable
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        dataTable.TableName = tableNamePrepared;
                        adapter.Fill(dataTable);

                        // Pass the data table to the view
                        return dataTable;
                    }
                }
            }
        }

        public Dictionary<string, int[]> RetrieveJson(string tableName, int id = 0)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                // Create a SQL command to fetch data from the specified table
                string query;
                if (id != 0) query = $"SELECT datasource FROM {tableName} WHERE ID = {id}";
                else query = $"SELECT datasource FROM {tableName} WHERE ID = (SELECT MAX(Id) From {tableName})";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var re = reader["datasource"];
                            var o = JsonConvert.DeserializeObject<Dictionary<string, int[]>>(re.ToString());
                            return o;

                        }
                    }
                    return null;
                }
            }
        }

        public void DeleteRowById(string tableName, int id = 0)
        {

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                string deleteSql;
                if (id != 0)
                {
                    deleteSql = $"DELETE FROM {tableName} WHERE Id = @Id";
                }
                else
                {
                    deleteSql = $"DELETE FROM {tableName} WHERE Id = (SELECT MAX(Id) From {tableName})";
                }

                using (SqlCommand command = new SqlCommand(deleteSql, connection))
                {

                    command.Parameters.AddWithValue("@Id", id);

                    int rowsAffected = command.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        Console.WriteLine($"Row with ID {id} deleted successfully from table {tableName}.");
                    }
                    else
                    {
                        Console.WriteLine($"Row with ID {id} not found in table {tableName}.");
                    }
                }
            }
        }

        public void UpdateTable(string tableName, Dictionary<string, object> keyValuePairs)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                // Create a SQL command to update data in the specified table
                string query = $"UPDATE {tableName} SET ";

                SqlCommand command = new SqlCommand();
                foreach (var kvp in keyValuePairs)
                {
                    if (kvp.Key == "ID") continue;

                    // Append column names and parameters to the query
                    query += $"{kvp.Key} = @{kvp.Key}, ";

                    // Add parameters to the SqlCommand

                    command.Parameters.AddWithValue("@" + kvp.Key, kvp.Value.ToString());

                }

                // Remove the trailing comma and space
                query = query.TrimEnd(',', ' ');

                // Append the WHERE clause (assuming "ID" is always present)
                query += " WHERE ID = @ID";
                command.Parameters.AddWithValue("@ID", keyValuePairs["ID"]);

                // Set the connection and query for the command
                command.Connection = connection;
                command.CommandText = query;

                // Execute the update query
                command.ExecuteNonQuery();
            }
        }
       
        public void CreateMetadataTable(List<Title> tables)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Metadata' and xtype='U')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"CREATE TABLE Metadata (");
            sb.AppendLine($"[TableName]  NVARCHAR(MAX),");
            sb.AppendLine($"[SheetName]  NVARCHAR(MAX)");
            sb.AppendLine(")");
            sb.AppendLine("END");
            ExecCommand(sb.ToString());
            InsertMetadata(tables);
        }

        public void InsertMetadata(List<Title> tables)
        {
            DataTable tbl = new DataTable();
            tbl.Columns.Add(new DataColumn("TableName"));
            tbl.Columns.Add(new DataColumn("SheetName"));

            foreach (Title table in tables)
            {
                DataRow dr = tbl.NewRow();
                dr["TableName"] = table.name;
                dr["SheetName"] = table.sheet;
                tbl.Rows.Add(dr);
                dr.AcceptChanges();
            }

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    objbulk.BulkCopyTimeout = 100;
                    objbulk.DestinationTableName = "Metadata";
                    objbulk.ColumnMappings.Add("SheetName", "SheetName");
                    objbulk.ColumnMappings.Add("TableName", "TableName");

                    try
                    {
                        objbulk.WriteToServer(tbl);
                        transaction.Commit();
                    }

                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        public void CreateConfigTableJson(Dictionary<string, string[,]> data)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ExcelConfig' and xtype='U')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"CREATE TABLE ExcelConfig (");
            sb.AppendLine($"Id INT PRIMARY KEY IDENTITY(1,1), [Config]  NVARCHAR(MAX)");
            sb.AppendLine(")");
            sb.AppendLine("END");
            ExecCommand(sb.ToString());
            InsertConfigTableJson(data);
        }

        public void InsertConfigTableJson(Dictionary<string, string[,]> data)
        {
            string dictionaryJson = JsonConvert.SerializeObject(data);

            DataTable tbl = new DataTable();

            tbl.Columns.Add(new DataColumn("Config"));

            DataRow dr = tbl.NewRow();
            if (!tbl.Columns.Contains("Config"))
            {
                tbl.Columns.Add(new DataColumn("Config"));
            }

            dr["Config"] = dictionaryJson;

            tbl.Rows.Add(dr);

            dr.AcceptChanges();

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    objbulk.BulkCopyTimeout = 100;
                    objbulk.DestinationTableName = "ExcelConfig";
                    objbulk.ColumnMappings.Add("Config", "Config");

                    try
                    {
                        objbulk.WriteToServer(tbl);
                        transaction.Commit();
                    }

                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }      

           
        public string GetJsonConfig()
        {

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                string sql = "SELECT Config FROM ExcelConfig WHERE Id = (SELECT MAX(Id) From ExcelConfig)";

                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    //command.Parameters.AddWithValue("Id", 1);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string xmlStringFromDb = reader["Config"].ToString();
                            return xmlStringFromDb;
                        }
                    }
                    return null;
                }
            }
        }       

        public void CreateTable()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{tableName}' and xtype='U')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"CREATE TABLE {tableName} (");
            sb.AppendLine($"ID INT PRIMARY KEY IDENTITY(1,1),");
            sb.AppendLine($"datasource NVARCHAR(255),");


            for (int colIdx = 0; colIdx < table.columnCount; colIdx++)
            {
                string type = CheckDataType(1, colIdx);
                string prefix = " ";
                if (colIdx != 0) prefix = ",";

                sb.AppendLine($"{prefix}[{table.columns[colIdx]}] {type}");

            }
            sb.AppendLine(")");
            sb.AppendLine("END");
            ExecCommand(sb.ToString());
        }


        public string CheckDataType(int i, int j)
        {
            Type type = table.values[i, j].GetType();

            if (type == typeof(int)) return "INT";
            else if (type == typeof(double)) return "FLOAT(53)";
            else if (type == typeof(float)) return "FLOAT(53)";
            else if (type == typeof(decimal)) return "DECIMAL(2, 2)";
            else if (type == typeof(DateTime)) return "DATETIME";
            else return "NVARCHAR(255)";
        }

        public void TableInsert()
        {
            DataTable tbl = new DataTable();

            for (int colIdx = 0; colIdx < table.columnCount; colIdx++)
            {
                Type type = table.values[1, colIdx].GetType();
                tbl.Columns.Add(new DataColumn(table.columns[colIdx], type));

            }
            tbl.Columns.Add(new DataColumn("datasource"));

            var rowCount = table.rowCount;
            for (int i = rowCount.Value - 1; i >= 0; i--)
            {
                DataRow dr = tbl.NewRow();


                var counter = 0;

                for (int j = 0; j < table.columnCount; j++)
                {
                    if (table.values[i, j] == null) dr[table.columns[j]] = DBNull.Value;
                    else dr[table.columns[j]] = table.values[i, j];
                    counter++;

                }
                var test = (Dictionary<string, int[]>)table.datasource[i, table.columnCount.Value - counter];
                string dictionaryJson = JsonConvert.SerializeObject(test);

                dr["datasource"] = dictionaryJson;

                tbl.Rows.Add(dr);
            }

            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    objbulk.DestinationTableName = tableName;

                    for (int j = 0; j < table.columnCount; j++)
                    {
                        objbulk.ColumnMappings.Add(table.columns[j], table.columns[j]);
                    }
                    objbulk.ColumnMappings.Add("datasource", "datasource");

                    try
                    {
                        objbulk.WriteToServer(tbl);
                        transaction.Commit();
                    }

                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        // 0 = DB name
        private const string createDbCmd = @"
        IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = '{0}')
        BEGIN
        CREATE DATABASE[{0}]
        END";

        public void CreateDb(string DbName)
        {
            var cmd = string.Format(createDbCmd, DbName);
            ExecCommand(cmd);
        }

        // Helper function to execute Sql commands
        private static void ExecCommand(string queryString)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
        }
    }
}

