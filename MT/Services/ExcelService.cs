using OfficeOpenXml;
using Table = MT.Models.Table;
using Column = MT.Models.Column;
using MT.Models;

namespace MT.Services
{
    public interface IExcelService
    {
        List<Table> FFinalTables { get; set; }
        List<Formula> formulas { get; set; }
        List<DataEntry> dataEntries { get; set; }
    }

    public class ExcelService
    {
        public ExcelWorksheets wss;
        public List<Title> tables = new List<Title>();
        private List<Column> columns = new List<Column>();
        public List<DataEntry> dataEntries = new List<DataEntry>();
        private int tableCounter = 0;
        private int nullCounter = 0;
        public List<Table> finalTables { get; set; }
        public List<Tableinfo> tableInfos = new List<Tableinfo>();
        public List<Formula> formulas = new List<Formula>();
        public List<dynamic> models = new List<dynamic>();

        private static string filePath = ".\\wwwroot\\Uploads\\Test.xlsx";


        public ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage excelPackage = new ExcelPackage(filePath);

            DeleteFile();

            wss = excelPackage.Workbook.Worksheets;

            var table = TableDetection();
            MapToTable(table);
            TableInfo(table);
            new DbService().CreateMetadataTable(table);


            foreach (Table finalTable in finalTables)
            {
                var dbService = new DbService(finalTable);
                //dbService.CreateDb("Test");
                dbService.CreateTable();
                dbService.TableInsert();
            }
        }

        public void DeleteFile()
        {
            try
            {
                // Check if the file exists before attempting to delete it
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public List<Title> TableDetection()
        {
            var db = new DbService();

            foreach (var ws in wss)
            {
                tableCounter++;
                int colCount;
                int rowCount;

                try
                {
                    //get column count of document
                    colCount = ws.Dimension.End.Column;
                    //get row count of document
                    rowCount = ws.Dimension.End.Row;
                }

                catch
                {
                    colCount = 0;
                    rowCount = 0;
                }

                if (ws.Name == "Formulas")
                {
                    List<Formula> formulas = new List<Formula>();
                    for (int row = rowCount; row >= 2; row--)
                    {
                        if (ws.Cells[row, 1].Formula == "" && ws.Cells[row, 1].Value == null) continue;
                        var newFormula = new Formula();
                        newFormula.formula = ws.Cells[row, 1].Formula ?? null;
                        newFormula.result = ws.Cells[row, 1].Value?.ToString() ?? null;
                        newFormula.context = ws.Cells[row, 2].Value?.ToString();
                        newFormula.i = row - 1;
                        formulas.Add(newFormula);
                    }
                    db.CreateFormulaTable(formulas);

                    string[,] valuesFormulas = new string[rowCount, colCount];

                    for (int row = rowCount; row >= 1; row--)
                    {
                        for (int col = colCount; col >= 1; col--)
                        {
                            if (ws.Cells[row, col].Value == null) valuesFormulas[row - 1, col - 1] = "";
                            else if (ws.Cells[row, col].Formula == "") valuesFormulas[row - 1, col - 1] = ws.Cells[row, col].Value.ToString();
                            else valuesFormulas[row - 1, col - 1] = "=" + ws.Cells[row, col].Formula;
                        }
                    }
                    continue;
                }

                string[,] values = new string[rowCount, colCount];

                for (int row = rowCount; row >= 1; row--)
                {
                    for (int col = colCount; col >= 1; col--)
                    {
                        values[row - 1, col - 1] = TableRecognition(row, col, ws);
                    }
                }
            }

            return tables;
        }

        public void TableInfo(List<Title> tables)
        {
            foreach (Title table in tables)
            {
                var newInfo = new Tableinfo();
                newInfo.name = table.name;
                newInfo.columnCount = table.columns.Count;
                newInfo.rowCount = table.columns.LastOrDefault().entries.Count;
                tableInfos.Add(newInfo);
            }
        }

        public string TableRecognition(int i, int j, ExcelWorksheet ews)
        {
            if (ews.Cells[i, j].Value == null)
            {
                nullCounter++;
                if (nullCounter == 5 * ews.Dimension.End.Column)
                {
                    tableCounter++;
                }
                return "";
            }

            else if (IsDataCell(i, j, ews) || ews.Cells[i, j].Value is "n/a" or "N/A")
            {
                // Add new entry (N/A)
                if (ews.Cells[i, j].Value is "n/a" or "N/A")
                {
                    var newEntry = new DataEntry();
                    newEntry.i = i;
                    newEntry.j = j;
                    newEntry.value = null;
                    dataEntries.Add(newEntry);
                    newEntry.tableCount = tableCounter;
                }

                // Add new entry
                else
                {
                    Dictionary<string, int[]> dic = new Dictionary<string, int[]>();
                    int[] myArray = new int[2] { i, j };
                    dic.Add(ews.Name, myArray);
                    var newEntry = new DataEntry();
                    newEntry.datasource = dic;
                    newEntry.i = i;
                    newEntry.j = j;
                    newEntry.value = ews.Cells[i, j].Value;
                    dataEntries.Add(newEntry);
                    newEntry.tableCount = tableCounter;
                }
            }

            // Add new header
            else if (IsHeaderCell(i, j, ews))
            {
                var column = new Column();
                column.tableCount = tableCounter;
                column.count = j;
                column.Name = ews.Cells[i, j].Value.ToString();
                column.entries = dataEntries.Where(p => p.j == j && p.tableCount == tableCounter).ToList();
                columns.Add(column);
            }

            // Add new title
            else if (IsTitleCell(i, j, ews))
            {
                var title = new Title();
                title.tableCount = tableCounter;
                title.name = ews.Cells[i, j].Value.ToString();
                title.columns = columns.Where(C => C.tableCount == tableCounter).ToList();
                title.sheet = ews.Name;
                tableCounter++;
                tables.Add(title);
            }

            nullCounter = 0;

            return ews?.Cells[i, j]?.Value?.ToString();
        }

        public bool IsTitleCell(int i, int j, ExcelWorksheet ews)
        {
            // Cell is type of string and there is a seperator
            if (ews.Cells[i, j].Value.GetType() == typeof(string) && columns.Where(c => c.tableCount == tableCounter).Count() != 0 && ews.Cells[i + 1, j].Value == null)
            {
                return true;
            }

            // A a title cell has already been found and in the last row scanned, a title cell has been detected
            else if (j != 1 && tables.Where(t => t.tableCount == tableCounter).Count() != 0 && (IsTitleCell(i + 1, j, ews) || IsTitleCell(i + 1, j - 1, ews) || IsTitleCell(i + 1, j + 1, ews)))
            {
                return true;
            }

            // A header cell has been found, C[i,j-1] is empty, C[i,j+1] empty and j is the table’s first column
            else if (j != 1 && columns.Where(c => c.tableCount == tableCounter).Count() != 0 && ews.Cells[i, j + 1].Value == null && ews.Cells[i, j - 1].Value == null && columns.Where(c => c.tableCount == tableCounter).First().count == columns.Where(c => c.tableCount == tableCounter).MinBy(c => c.count).count)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool IsHeaderCell(int i, int j, ExcelWorksheet ews)
        {
            if (ews.Cells[i, j].Value == null) return false;

            // Cell has type string and row below is not null
            else if (columns.Where(c => c.tableCount == tableCounter).Count() != 0 && ews.Cells[i, j].Value.GetType() == typeof(string) && ews.Cells[i + 1, j].Value != null)
            {
                return true;
            }

            // Cell or neighbour cells have borders and row below is not null
            else if (j != 1 && HasBorders(i, j, ews) && (HasBorders(i, j + 1, ews) || j != 0 && HasBorders(i, j - 1, ews)) && ews.Cells[i + 1, j].Value != null)
            {
                return true;
            }

            // Cell and neighbour cells have borders and row below is not null
            else if (j != 1 && HaveSimilarFormat(i, j, i, j - 1, ews) && HaveSimilarFormat(i, j, i, j + 1, ews) && ews.Cells[i + 1, j].Value != null)
            {
                return true;
            }

            // Rules for similar formatting between two consecutive cells and to distinguish headers from data entries
            else if (j != 1 && ews.Cells[i + 1, j].Value != null && ews.Cells[i + 1, j].Value.GetType() != typeof(string) && ews.Cells[i, j].Value.GetType() == typeof(string) && !HaveSimilarFormat(i, j, i + 1, j, ews) && ews.Cells[i, j - 1].Value != null && HaveSimilarFormat(i, j, i, j - 1, ews))
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool IsDataCell(int i, int j, ExcelWorksheet ews)
        {
            if (i != 1 && HaveSimilarFormat(i, j, i - 1, j, ews) && HaveSimilarFormat(i, j, i + 1, j, ews))
            {
                return true;
            }

            else if (i != 1 && HaveSimilarFormat(i, j, i - 1, j, ews) && ews.Cells[i + 1, j].Value == null)
            {
                return true;
            }

            else if (i != 1 && HaveSimilarFormat(i, j, i + 1, j, ews) && IsHeaderCell(i - 1, j, ews) == true)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool HaveSimilarFormat(int i, int j, int ii, int jj, ExcelWorksheet ws)
        {
            if (ii == 0 || jj == 0 || ws.Cells[ii, jj].Value == null) return false;


            else if (ws.Cells[i, j].Style.Font.Size == ws.Cells[ii, jj].Style.Font.Size && ws.Cells[i, j].Style.Font.Color.Tint == ws.Cells[ii, jj].Style.Font.Color.Tint)
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        public bool HasBorders(int i, int j, ExcelWorksheet ws)
        {
            if (i == 0 || j == 0 || ws.Cells[i, j].Value == null) return false;

            else if (ws.Cells[i, j].Style.Border.Top != null || ws.Cells[i, j].Style.Border.Bottom != null || ws.Cells[i, j].Style.Border.Left != null || ws.Cells[i, j].Style.Border.Right != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        // Mapping from column-based model to row-based
        public List<Table> MapToTable(List<Title> tableNames)
        {
            List<Table> finalTables = new List<Table>();

            foreach (Title t in tableNames)
            {

                int colCount = t.columns.Count;
                int rowCount = t.columns.Select(x => x.entries.Count).Max();

                Table table = new Table();
                table.tableName = t.name;
                table.columns = new string[colCount];
                table.values = new object[rowCount, colCount];
                table.datasource = new object[rowCount, colCount];
                table.columnCount = colCount;
                table.rowCount = rowCount;

                var columns = t.columns.OrderBy(x => x.count).ToList();
                for (int colIdx = 0; colIdx < columns.Count; colIdx++)
                {
                    var col = columns[colIdx];
                    table.columns[colIdx] = col.Name;

                    for (int rowIdx = 0; rowIdx < col.entries.Count; rowIdx++)
                    {
                        var entry = col.entries[rowIdx];
                        table.values[rowIdx, colIdx] = entry.value;
                        table.datasource[rowIdx, colIdx] = entry.datasource;
                    }
                }
                finalTables.Add(table);
            }

            this.finalTables = finalTables;
            return finalTables;
        }
    }
}



