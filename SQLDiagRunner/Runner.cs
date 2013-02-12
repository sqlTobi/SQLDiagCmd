using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

using OfficeOpenXml;
using System.Diagnostics;

namespace SQLDiagRunner
{
    public class Runner
    {
        private readonly HashSet<string> _dictWorksheet = new HashSet<string>();

        public void ExecuteQueries
        (
            string servername,
            string username,
            string password,
            string scriptLocation,
            string outputFolder,
            IList<string> databases,
            bool useTrusted,
            bool autoFitColumns,
            int queryTimeoutSeconds
        )
        {
            _dictWorksheet.Clear();

            string dateString = DateTime.Now.ToString("yyyyMMdd_hhmmss_");
            List<SqlQuery> queries = new List<SqlQuery>();

            foreach (var file in Directory.EnumerateFiles(scriptLocation, "*.sql"))
            {
                var parser = new QueryFileParser(file);
                queries.AddRange(parser.Load());
            }

            string outputFilepath = GetOutputFilepath(outputFolder, servername, dateString);

            using (var fs = new FileStream(outputFilepath, FileMode.Create))
            {
                using (var pck = new ExcelPackage(fs))
                {
                    string connectionString = GetConnectionStringTemplate(servername, "master", username, password, useTrusted);
                    var serverQueries = queries.Where(q => q.ServerScope).ToList();
                    ExecuteQueriesAndSaveToExcel(pck, connectionString, serverQueries, "", "", autoFitColumns, queryTimeoutSeconds);

                    if (databases.Count > 0)
                    {
                        int databaseNo = 1;
                        var dbQueries = queries.Where(q => !q.ServerScope).ToList();
                        foreach (var db in databases)
                        {
                            connectionString = GetConnectionStringTemplate(servername, db, username, password, useTrusted);
                            ExecuteQueriesAndSaveToExcel(pck, connectionString, dbQueries, db.Trim(),
                                                         databaseNo.ToString(CultureInfo.InvariantCulture),
                                                         autoFitColumns, queryTimeoutSeconds);
                            databaseNo++;
                        }
                    }

                    foreach (var file in Directory.EnumerateFiles(scriptLocation, "*.ps1"))
                    {
                        ExecutePoShAndSaveToExcel(pck, file, autoFitColumns);
                    }

                    pck.Save();
                }
            }
        }

        private static string GetOutputFilepath(string outputFolder, string servername, string dateString)
        {
            string ret = Directory.Exists(outputFolder)
                             ? Path.Combine(outputFolder, dateString + servername.ReplaceInvalidFilenameChars("_") + ".xlsx")
                             : outputFolder;

            return ret;
        }

        private void ExecuteQueriesAndSaveToExcel
        (
            ExcelPackage pck,
            string connectionstring,
            IEnumerable<SqlQuery> queries,
            string database,
            string worksheetPrefix,
            bool autoFitColumns,
            int queryTimeoutSeconds
        )
        {
            foreach (var q in queries)
            {
                string query = GetQueryText(q, database);
                string worksheetName = GetWorkSheetName(q.Title, worksheetPrefix);
                string cell = "A1";

                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(worksheetName);
                try
                {
                    if (!String.IsNullOrEmpty(database))
                    {
                        ws.Cells[cell].Value = database;
                        cell = "A2";
                    }

                    DataTable datatable = QueryExecutor.Execute(connectionstring, query, queryTimeoutSeconds);

                    if (datatable.Rows.Count > 0)
                    {
                        ExcelRangeBase range = ws.Cells[cell].LoadFromDataTable(datatable, true);

                        ws.Row(Int16.Parse(cell.Substring(1))).Style.Font.Bold = true;

                        // find datetime columns and set formatting
                        int numcols = datatable.Columns.Count;
                        for (int i = 0; i < numcols; i++)
                        {
                            var column = datatable.Columns[i];
                            if (column.DataType == typeof(DateTime))
                            {
                                ws.Column(i + 1).Style.Numberformat.Format = "yyyy-mm-dd hh:MM:ss";
                            }
                        }

                        if (autoFitColumns)
                        {
                            range.AutoFitColumns();
                        }
                    }
                    else
                    {
                        ws.Cells[cell].Value = "None Found";
                    }
                }
                catch (Exception ex)
                {
                    ws.Cells[cell].Value = ex.Message;
                }
            }
        }

        private void ExecutePoShAndSaveToExcel(ExcelPackage pck, string ps1Path, bool autoFitColumns)
        {
            string worksheetName = String.Empty;
            string cell = "A1";
            
            try
            {
                DataTable datatable = RunPowershellScript(ps1Path);
                worksheetName = GetWorkSheetName(datatable.TableName, String.Empty);
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(worksheetName);

                if (datatable.Rows.Count > 0)
                {
                    ExcelRangeBase range = ws.Cells[cell].LoadFromDataTable(datatable, true);
                    ws.Row(Int16.Parse(cell.Substring(1))).Style.Font.Bold = true;

                    // find datetime columns and set formatting
                    int numcols = datatable.Columns.Count;
                    for (int i = 0; i < numcols; i++)
                    {
                        var column = datatable.Columns[i];
                        if (column.DataType == typeof(DateTime))
                        {
                            ws.Column(i + 1).Style.Numberformat.Format = "yyyy-mm-dd hh:MM:ss";
                        }
                    }

                    if (autoFitColumns)
                    {
                        range.AutoFitColumns();
                    }
                }
                else
                {
                    ws.Cells[cell].Value = "None Found";
                }
            }
            catch (Exception ex)
            {
                //ws.Cells[cell].Value = ex.Message;
            }
        }


        private string GetQueryText(SqlQuery q, string database)
        {
            return string.IsNullOrEmpty(database) ? q.QueryText : "USE [" + database + "] \r\n " + q.QueryText;
        }

        private string GetWorkSheetName(string queryTitle, string worksheetPrefix)
        {
            string worksheetName = string.IsNullOrEmpty(worksheetPrefix) ? queryTitle : worksheetPrefix + " " + queryTitle;

            string worksheetNameSanitised = SanitiseWorkSheetName(worksheetName);

            // Check if name already exists: 31 char limit for worksheet names!
            while (_dictWorksheet.Contains(worksheetNameSanitised))
            {
                worksheetNameSanitised = worksheetNameSanitised.RandomiseLastNChars(3);
            }

            _dictWorksheet.Add(worksheetNameSanitised);

            return worksheetNameSanitised;
        }

        private string SanitiseWorkSheetName(string wsname)
        {
            var s = wsname.RemoveInvalidExcelChars();

            return s.Substring(0, Math.Min(31, s.Length));
        }

        private string GetConnectionStringTemplate
        (
            string servername,
            string database,
            string username,
            string password,
            bool trusted
        )
        {
            const string trustedConnectionStringTemplate = "server={0};database={1};trusted_Connection=True";
            const string sqlLoginConnectionStringTemplate = "server={0};database={1};username={2};password={3}";

            if (trusted)
            {
                return string.Format(trustedConnectionStringTemplate, servername, database);
            }

            return string.Format(sqlLoginConnectionStringTemplate, servername, database, username, password);

        }

        private static DataTable RunPowershellScript(string ps1Path)
        {
            DataTable dt = new DataTable();

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"powershell.exe";
            startInfo.Arguments = String.Format(@"& '{0}'", ps1Path);
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();

            string errors = process.StandardError.ReadToEnd();
            if (!String.IsNullOrEmpty(errors))
            {
                dt.Columns.Add("ErrorMessage", typeof(string));
                DataRow row = dt.NewRow();
                row["ErrorMessage"] = errors;
                dt.Rows.Add(row);
            }
            else
            {
                string output = process.StandardOutput.ReadToEnd();
                if (!String.IsNullOrEmpty(output))
                {
                    try
                    {
                        using (TextReader sr = new StringReader(output.Replace(System.Environment.NewLine, String.Empty)))
                            dt.ReadXml(sr);
                    }
                    catch (Exception ex)
                    {
                        dt.Columns.Add("ErrorMessage", typeof(string));
                        DataRow row = dt.NewRow();
                        row["ErrorMessage"] = ex.Message;
                        dt.Rows.Add(row);
                    }
                }
                else
                {
                    dt.Columns.Add("ErrorMessage", typeof(string));
                    DataRow row = dt.NewRow();
                    row["ErrorMessage"] = "Powershell script has created no output.";
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }
    }

}

